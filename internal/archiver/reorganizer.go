package archiver

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"go.uber.org/zap"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/outlook"
	"outlook-archiver/internal/router"
	"outlook-archiver/pkg/comutil"
)

// ReorganizeResult 全量整理结果
type ReorganizeResult struct {
	Phase1 ArchiveResult // 阶段一：邮箱归档
	Phase2 RectifyResult // 阶段二：PST 纠偏
}

// RectifyResult PST 纠偏结果
type RectifyResult struct {
	OurPSTScanned    int           // 本工具 PST 扫描数
	LegacyPSTScanned int           // 第三方 PST 扫描数
	TotalRectified   int           // 纠偏邮件数
	TotalMigrated    int           // 第三方 PST 迁移数
	TotalFailed      int
	Duration         time.Duration
}

// Reorganizer 全量整理引擎
type Reorganizer struct {
	cfg      *config.Config
	bridge   *outlook.COMBridge
	archiver *Archiver
	logger   *zap.Logger
}

func NewReorganizer(cfg *config.Config, bridge *outlook.COMBridge, archiver *Archiver, logger *zap.Logger) *Reorganizer {
	return &Reorganizer{
		cfg:      cfg,
		bridge:   bridge,
		archiver: archiver,
		logger:   logger,
	}
}

// Reorganize 执行全量整理
func (r *Reorganizer) Reorganize(ctx context.Context) (*ReorganizeResult, error) {
	result := &ReorganizeResult{}

	// 阶段一：邮箱全量归档
	r.logger.Info("开始全量整理 - 阶段一：邮箱归档")
	phase1Opts := ArchiveOptions{
		MaxBatchSize: 0, // 不限制批次
		SafeDelay:    0, // 不延迟
		MoveInterval: time.Duration(r.cfg.MoveIntervalMs) * time.Millisecond,
		DryRun:       r.cfg.DryRun,
	}
	phase1Res, err := r.archiver.Archive(ctx, phase1Opts)
	if err != nil {
		r.logger.Error("阶段一归档失败", zap.Error(err))
		return nil, fmt.Errorf("phase 1 failed: %w", err)
	}
	if phase1Res != nil {
		result.Phase1 = *phase1Res
	}

	// 阶段二：PST 纠偏
	r.logger.Info("开始全量整理 - 阶段二：PST 纠偏")
	startPhase2 := time.Now()

	ourPSTs, legacyPSTs, err := r.discoverPSTs()
	if err != nil {
		r.logger.Error("发现 PST 文件失败", zap.Error(err))
		return nil, fmt.Errorf("failed to discover PSTs: %w", err)
	}

	result.Phase2.OurPSTScanned = len(ourPSTs)
	result.Phase2.LegacyPSTScanned = len(legacyPSTs)

	r.logger.Info("PST 扫描结果", zap.Int("our_psts", len(ourPSTs)), zap.Int("legacy_psts", len(legacyPSTs)))

	// 处理本工具的 PST 纠偏
	for _, pstPath := range ourPSTs {
		if ctx.Err() != nil {
			return result, ctx.Err()
		}
		r.rectifyOurPST(ctx, pstPath, &result.Phase2)
	}

	// 处理第三方 PST 的全量迁移
	for _, pstPath := range legacyPSTs {
		if ctx.Err() != nil {
			return result, ctx.Err()
		}
		r.migrateLegacyPST(ctx, pstPath, &result.Phase2)
	}

	result.Phase2.Duration = time.Since(startPhase2)
	r.logger.Info("全量整理完成", zap.Duration("duration", result.Phase2.Duration))

	return result, nil
}

// discoverPSTs 发现所有的 PST 文件并分类
func (r *Reorganizer) discoverPSTs() (ourPSTs []string, legacyPSTs []string, err error) {
	// 扫描 pst_root_path
	entries, err := os.ReadDir(r.cfg.PSTRootPath)
	if err == nil {
		for _, entry := range entries {
			if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), ".pst") {
				path := filepath.Join(r.cfg.PSTRootPath, entry.Name())
				if router.IsOurPSTName(entry.Name()) {
					ourPSTs = append(ourPSTs, path)
				} else {
					legacyPSTs = append(legacyPSTs, path)
				}
			}
		}
	} else if !os.IsNotExist(err) {
		r.logger.Warn("扫描 pst_root_path 失败", zap.Error(err))
	}

	// 扫描 legacy_pst_scan_paths
	legacyPaths, err := DiscoverLegacyPSTs(r.cfg.LegacyPSTScanPaths, r.cfg.PSTRootPath)
	if err != nil {
		r.logger.Warn("扫描 legacy_pst_scan_paths 失败", zap.Error(err))
	}
	legacyPSTs = append(legacyPSTs, legacyPaths...)

	return ourPSTs, legacyPSTs, nil
}

// DiscoverLegacyPSTs 自动发现第三方 PST 文件
func DiscoverLegacyPSTs(scanPaths []string, ourRootPath string) ([]string, error) {
	var legacyPSTs []string
	
	// 确保 ourRootPath 转换为绝对路径便于比较
	absOurRoot, err := filepath.Abs(ourRootPath)
	if err != nil {
		absOurRoot = ourRootPath
	}

	for _, scanPath := range scanPaths {
		absScanPath, err := filepath.Abs(scanPath)
		if err != nil {
			continue
		}
		
		// 如果扫描目录和我们的根目录一致，跳过（已经在上面处理过）
		if strings.EqualFold(absScanPath, absOurRoot) {
			continue
		}

		entries, err := os.ReadDir(scanPath)
		if err != nil {
			continue
		}

		for _, entry := range entries {
			if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), ".pst") {
				if !router.IsOurPSTName(entry.Name()) {
					legacyPSTs = append(legacyPSTs, filepath.Join(scanPath, entry.Name()))
				}
			}
		}
	}
	return legacyPSTs, nil
}

func (r *Reorganizer) rectifyOurPST(ctx context.Context, pstPath string, res *RectifyResult) {
	r.logger.Info("正在纠偏 PST", zap.String("path", pstPath))
	r.processPST(ctx, pstPath, res, false)
}

func (r *Reorganizer) migrateLegacyPST(ctx context.Context, pstPath string, res *RectifyResult) {
	r.logger.Info("正在迁移第三方 PST", zap.String("path", pstPath))
	r.processPST(ctx, pstPath, res, true)
}

func (r *Reorganizer) processPST(ctx context.Context, pstPath string, res *RectifyResult, forceMigrate bool) {
	// 挂载 PST
	pstRoot, err := r.bridge.EnsurePSTMountedByPath(pstPath)
	if err != nil {
		r.logger.Error("挂载 PST 失败", zap.String("path", pstPath), zap.Error(err))
		return
	}
	defer comutil.SafeRelease(pstRoot)

	// 遍历文件夹
	folders, err := r.bridge.WalkPSTFolders(pstRoot)
	if err != nil {
		r.logger.Error("遍历 PST 文件夹失败", zap.String("path", pstPath), zap.Error(err))
		return
	}

	currentPSTName := filepath.Base(pstPath)

	for _, folder := range folders {
		if ctx.Err() != nil {
			return
		}
		r.processPSTFolder(ctx, folder, currentPSTName, res, forceMigrate)
	}
}

func (r *Reorganizer) processPSTFolder(ctx context.Context, folder outlook.FolderInfo, currentPSTName string, res *RectifyResult, forceMigrate bool) {
	itemsVar, err := comutil.SafeGetProperty(folder.Dispatch, "Items")
	if err != nil {
		return
	}
	items := itemsVar.ToIDispatch()
	if items == nil {
		return
	}
	defer comutil.SafeRelease(items)

	// 仅筛选 IPM.Note 类型的邮件
	filter := "[MessageClass] = 'IPM.Note'"
	restrictedVar, err := comutil.SafeCallMethod(items, "Restrict", filter)
	if err != nil {
		return
	}
	restricted := restrictedVar.ToIDispatch()
	if restricted == nil {
		return
	}
	defer comutil.SafeRelease(restricted)

	countVar, err := comutil.SafeGetProperty(restricted, "Count")
	if err != nil {
		return
	}
	count := 0
	if countVar.Value() != nil {
		switch v := countVar.Value().(type) {
		case int32:
			count = int(v)
		case int:
			count = v
		case int16:
			count = int(v)
		case float64:
			count = int(v)
		case int64:
			count = int(v)
		}
	}

	for i := count; i >= 1; i-- {
		if ctx.Err() != nil {
			break
		}

		itemVar, err := comutil.SafeCallMethod(restricted, "Item", i)
		if err != nil {
			res.TotalFailed++
			continue
		}
		item := itemVar.ToIDispatch()
		if item == nil {
			res.TotalFailed++
			continue
		}
		itemRef := comutil.NewCOMRef(item, fmt.Sprintf("mail_%d", i))

		subject := r.bridge.GetSubject(itemRef.Dispatch())
		mailTime, err := r.bridge.GetMailTime(itemRef.Dispatch(), folder.TimeField)
		if err != nil {
			r.logger.Warn("获取邮件时间失败", zap.String("subject", subject), zap.Error(err))
			itemRef.Release()
			res.TotalFailed++
			continue
		}

		quarter := router.CalcQuarter(mailTime)
		targetPSTName := quarter.PSTFileName()

		// 判断是否需要移动
		needsMove := false
		if forceMigrate {
			needsMove = true // 强制迁移
		} else {
			if !strings.EqualFold(currentPSTName, targetPSTName) {
				needsMove = true // 本工具PST内时间纠偏
			}
		}

		if !needsMove {
			itemRef.Release()
			continue
		}

		if r.cfg.DryRun {
			r.logger.Info("[DRY RUN] 会进行纠偏/迁移",
				zap.String("subject", subject),
				zap.String("source_pst", currentPSTName),
				zap.String("target_pst", targetPSTName),
			)
			if forceMigrate {
				res.TotalMigrated++
			} else {
				res.TotalRectified++
			}
			itemRef.Release()
			time.Sleep(time.Duration(r.cfg.MoveIntervalMs) * time.Millisecond)
			continue
		}

		// 执行移动
		targetPSTRoot, err := r.bridge.EnsurePSTMounted(quarter, r.cfg.PSTRootPath)
		if err != nil {
			r.logger.Error("挂载目标 PST 失败", zap.Error(err))
			itemRef.Release()
			res.TotalFailed++
			continue
		}
		
		targetFolder, err := r.bridge.EnsurePSTFolder(targetPSTRoot, folder.FullPath)
		if err != nil {
			r.logger.Error("创建目标 PST 文件夹失败", zap.Error(err))
			comutil.SafeRelease(targetPSTRoot)
			itemRef.Release()
			res.TotalFailed++
			continue
		}

		err = r.bridge.MoveItem(itemRef.Dispatch(), targetFolder)
		if err != nil {
			r.logger.Error("移动纠偏邮件失败", zap.String("subject", subject), zap.Error(err))
			res.TotalFailed++
		} else {
			if forceMigrate {
				res.TotalMigrated++
			} else {
				res.TotalRectified++
			}
		}

		comutil.SafeRelease(targetFolder)
		comutil.SafeRelease(targetPSTRoot)
		itemRef.Release()

		time.Sleep(time.Duration(r.cfg.MoveIntervalMs) * time.Millisecond)
	}
}
