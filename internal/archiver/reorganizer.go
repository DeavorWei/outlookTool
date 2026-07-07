package archiver

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"go.uber.org/zap"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/monitor"
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
	OurPSTScanned    int // 本工具 PST 扫描数
	LegacyPSTScanned int // 第三方 PST 扫描数
	TotalRectified   int // 纠偏邮件数
	TotalMigrated    int // 第三方 PST 迁移数
	TotalFailed      int
	Duration         time.Duration
}

// Reorganizer 全量整理引擎
type Reorganizer struct {
	cfgProvider func() config.Config // 实时获取当前生效配置（支持热重载）
	bridge      *outlook.COMBridge
	archiver    *Archiver
	logger      *zap.Logger
}

// getCfg 返回当前生效配置的最新副本（非启动时快照），保证热重载后整理逻辑使用新配置
func (r *Reorganizer) getCfg() config.Config {
	if r.cfgProvider == nil {
		return config.Config{}
	}
	return r.cfgProvider()
}

// checkPSTSize 检查目标 PST 文件大小，接近/超过阈值时输出告警日志（需求 §5.3）
func (r *Reorganizer) checkPSTSize(quarter router.QuarterInfo) {
	path := quarter.PSTFilePath(r.getCfg().PSTRootPath)
	status, err := monitor.CheckPSTSize(path)
	if err != nil {
		r.logger.Warn("检查 PST 大小失败", zap.String("path", path), zap.Error(err))
		return
	}
	switch status {
	case monitor.PSTSizeWarning:
		r.logger.Warn("单 PST 文件接近 20GB 上限，建议尽快拆分或迁移", zap.String("path", path), zap.Int64("size_gb_threshold", 15))
	case monitor.PSTSizeCritical:
		r.logger.Error("单 PST 文件超过 20GB，存在损坏风险，请立即拆分或迁移", zap.String("path", path), zap.Int64("size_gb_threshold", 20))
	}
}

func NewReorganizer(cfgProvider func() config.Config, bridge *outlook.COMBridge, archiver *Archiver, logger *zap.Logger) *Reorganizer {
	return &Reorganizer{
		cfgProvider: cfgProvider,
		bridge:      bridge,
		archiver:    archiver,
		logger:      logger,
	}
}

// Reorganize 执行全量整理
func (r *Reorganizer) Reorganize(ctx context.Context, progressCb func(phase, processed, rectified int, currentPST string)) (*ReorganizeResult, error) {
	result := &ReorganizeResult{}
	cfg := r.getCfg()

	// 阶段一：邮箱全量归档
	r.logger.Info("开始全量整理 - 阶段一：邮箱归档")
	if progressCb != nil {
		progressCb(1, 0, 0, "")
	}
	phase1Opts := ArchiveOptions{
		MaxBatchSize: 0,                // 不限制批次
		SafeDelay:    0,                // 不延迟
		RetainDays:   cfg.RetainDays,   // 全量整理阶段一（主邮箱归档）依然保留最近 N 天
		MoveInterval: time.Duration(cfg.MoveIntervalMs) * time.Millisecond,
		DryRun:       cfg.DryRun,
		CopyOnly:     cfg.CopyOnly,
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
		r.rectifyOurPST(ctx, pstPath, &result.Phase2, progressCb)
	}

	// 处理第三方 PST 的全量迁移
	for _, pstPath := range legacyPSTs {
		if ctx.Err() != nil {
			return result, ctx.Err()
		}
		r.migrateLegacyPST(ctx, pstPath, &result.Phase2, progressCb)
	}

	result.Phase2.Duration = time.Since(startPhase2)
	r.logger.Info("全量整理完成", zap.Duration("duration", result.Phase2.Duration))

	return result, nil
}

// discoverPSTs 发现所有的 PST 文件并分类
func (r *Reorganizer) discoverPSTs() (ourPSTs []string, legacyPSTs []string, err error) {
	cfg := r.getCfg()
	// 扫描 pst_root_path
	entries, err := os.ReadDir(cfg.PSTRootPath)
	if err == nil {
		for _, entry := range entries {
			if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), ".pst") {
				path := filepath.Join(cfg.PSTRootPath, entry.Name())
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

	// 构造要扫描的历史目录列表
	scanPaths := make([]string, len(cfg.LegacyPSTScanPaths))
	copy(scanPaths, cfg.LegacyPSTScanPaths)

	if cfg.IncludeMountedPSTs {
		mountedPSTs, err := r.bridge.GetMountedPSTs()
		if err != nil {
			r.logger.Warn("获取已挂载的数据文件失败", zap.Error(err))
		} else {
			dirMap := make(map[string]bool)
			for _, pstPath := range mountedPSTs {
				dir := filepath.Dir(pstPath)
				// 排除 PSTRootPath
				if !strings.EqualFold(dir, cfg.PSTRootPath) {
					dirMap[dir] = true
				}
			}
			for dir := range dirMap {
				// 避免与现有的 scanPaths 重复
				exists := false
				for _, sp := range scanPaths {
					if strings.EqualFold(sp, dir) {
						exists = true
						break
					}
				}
				if !exists {
					scanPaths = append(scanPaths, dir)
				}
			}
		}
	}

	// 扫描 legacy_pst_scan_paths + mounted directories
	legacyPaths, err := DiscoverLegacyPSTs(scanPaths, cfg.PSTRootPath)
	if err != nil {
		r.logger.Warn("扫描额外 PST 目录失败", zap.Error(err))
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

		err = filepath.Walk(scanPath, func(path string, info os.FileInfo, err error) error {
			if err != nil {
				return nil // ignore errors
			}
			if !info.IsDir() && strings.HasSuffix(strings.ToLower(info.Name()), ".pst") {
				if !router.IsOurPSTName(info.Name()) {
					legacyPSTs = append(legacyPSTs, path)
				}
			}
			return nil
		})
		if err != nil {
			continue
		}
	}
	return legacyPSTs, nil
}

func (r *Reorganizer) rectifyOurPST(ctx context.Context, pstPath string, res *RectifyResult, progressCb func(phase, processed, rectified int, currentPST string)) {
	r.logger.Info("正在纠偏 PST", zap.String("path", pstPath))
	r.processPST(ctx, pstPath, res, false, progressCb)
}

func (r *Reorganizer) migrateLegacyPST(ctx context.Context, pstPath string, res *RectifyResult, progressCb func(phase, processed, rectified int, currentPST string)) {
	r.logger.Info("正在迁移第三方 PST", zap.String("path", pstPath))
	r.processPST(ctx, pstPath, res, true, progressCb)
}

func (r *Reorganizer) processPST(ctx context.Context, pstPath string, res *RectifyResult, forceMigrate bool, progressCb func(phase, processed, rectified int, currentPST string)) {
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
	// C4：WalkPSTFolders 返回带引用的文件夹切片，必须释放以避免 COM 引用泄漏
	defer func() {
		for _, f := range folders {
			if f.Dispatch != nil {
				comutil.SafeRelease(f.Dispatch)
			}
		}
	}()

	currentPSTName := filepath.Base(pstPath)

	for _, folder := range folders {
		if ctx.Err() != nil {
			return
		}
		_ = r.bridge.SubmitWithContext(ctx, func() error {
			r.processPSTFolder(ctx, folder, currentPSTName, res, forceMigrate, progressCb)
			return nil
		})
	}
}

func (r *Reorganizer) processPSTFolder(ctx context.Context, folder outlook.FolderInfo, currentPSTName string, res *RectifyResult, forceMigrate bool, progressCb func(phase, processed, rectified int, currentPST string)) {
	cfg := r.getCfg()
	itemsVar, err := comutil.SafeGetProperty(folder.Dispatch, "Items")
	if err != nil {
		return
	}
	defer itemsVar.Clear()
	items := itemsVar.ToIDispatch()
	if items == nil {
		return
	}
	defer comutil.SafeRelease(items)

	// 不再限制 IPM.Note 类型的邮件，获取所有类型对象
	items.AddRef()
	restricted := items
	defer comutil.SafeRelease(restricted)

	count, err := r.bridge.GetCount(restricted)
	if err != nil {
		r.logger.Warn("Failed to get count", zap.String("folder", folder.FullPath), zap.Error(err))
		return
	}
	if count == 0 {
		return
	}

	// M5：显式排序，保证倒序遍历语义正确
	if err := r.bridge.SortItems(restricted, folder.TimeField, true); err != nil {
		r.logger.Warn("排序集合失败，回退为索引顺序", zap.String("folder", folder.FullPath), zap.Error(err))
	}

	blockSize := cfg.StreamBlockSize
	if blockSize <= 0 {
		blockSize = 1000
	}

	pstCache := make(map[string]*ole.IDispatch)
	defer func() {
		for _, root := range pstCache {
			comutil.SafeRelease(root)
		}
	}()

	processed := 0 // 跨块累积的已处理计数，用于进度回调频率控制

	// M5：分块倒序处理，内存峰值由 blockSize 决定，与文件夹邮件总数解耦
	for end := count; end >= 1; end -= blockSize {
		if ctx.Err() != nil {
			break
		}

		start := end - blockSize + 1
		if start < 1 {
			start = 1
		}

		// 收集本块紧凑快照
		block := make([]compactMeta, 0, end-start+1)
		for i := start; i <= end; i++ {
			if ctx.Err() != nil {
				break
			}
			item, err := r.bridge.GetItem(restricted, i)
			if err != nil || item == nil {
				continue
			}
			entryID, errID := r.bridge.GetEntryID(item)
			if errID != nil || entryID == "" {
				comutil.SafeRelease(item)
				continue
			}
			subject := r.bridge.GetSubject(item)
			mailTime, errTime := r.bridge.GetMailTime(item, folder.TimeField)
			if errTime != nil {
				comutil.SafeRelease(item)
				continue
			}
			comutil.SafeRelease(item)
			block = append(block, compactMeta{
				entryID: entryID,
				subject: subject,
				ts:      mailTime.Unix(),
			})
		}

		// 块内倒序处理（保留 GetItemFromID 解耦 Move 的正确性设计：
		// Move 会从源文件夹移除 item，若直接在被遍历的 restricted 集合上操作会破坏游标；
		// GetItemFromID 拿到的 item 与 restricted 集合无引用关系，Move 后不影响后续索引访问。）
		for j := len(block) - 1; j >= 0; j-- {
			m := block[j]
			if ctx.Err() != nil {
				break
			}

			mailItem, err := r.bridge.GetItemFromID(m.entryID)
			if err != nil || mailItem == nil {
				r.logger.Warn("Failed to get item by EntryID", zap.String("subject", m.subject), zap.Error(err))
				res.TotalFailed++
				continue
			}
			itemRef := comutil.NewCOMRef(mailItem, "mail_"+m.entryID)

			mailTime := time.Unix(m.ts, 0)
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

			processed++

			if !needsMove {
				itemRef.Release()
				if progressCb != nil && processed%50 == 0 {
					progressCb(2, res.TotalRectified+res.TotalMigrated, res.TotalRectified, currentPSTName)
				}
				continue
			}

			cutoffTime := CalcCutoffTime(0, cfg.RetainDays)
			if !mailTime.Before(cutoffTime) {
				r.logger.Warn("近期邮件被强制归档",
					zap.String("subject", m.subject),
					zap.Time("mail_time", mailTime),
					zap.String("source_pst", currentPSTName),
					zap.String("target_pst", targetPSTName),
					zap.Int("retain_days", cfg.RetainDays),
				)
			}

			if cfg.DryRun {
				r.logger.Info("[DRY RUN] 会进行纠偏/迁移",
					zap.String("subject", m.subject),
					zap.String("source_pst", currentPSTName),
					zap.String("target_pst", targetPSTName),
				)
				if forceMigrate {
					res.TotalMigrated++
				} else {
					res.TotalRectified++
				}
				itemRef.Release()
				if progressCb != nil && processed%10 == 0 {
					progressCb(2, res.TotalRectified+res.TotalMigrated, res.TotalRectified, currentPSTName)
				}
				time.Sleep(time.Duration(cfg.MoveIntervalMs) * time.Millisecond)
				continue
			}

			// 执行移动
			targetPSTRoot, ok := pstCache[targetPSTName]
			if !ok {
				targetPSTRoot, err = r.bridge.EnsurePSTMounted(quarter, cfg.PSTRootPath)
				if err != nil {
					r.logger.Error("挂载目标 PST 失败", zap.Error(err))
					itemRef.Release()
					res.TotalFailed++
					continue
				}
				pstCache[targetPSTName] = targetPSTRoot
				r.checkPSTSize(quarter) // M12：挂载目标 PST 后检查大小并告警
			}

			targetFolder, err := r.bridge.EnsurePSTFolder(targetPSTRoot, folder.FullPath)
			if err != nil {
				r.logger.Error("创建目标 PST 文件夹失败", zap.Error(err))
				// 注意：targetPSTRoot 已存入 pstCache，由 defer 统一释放，此处不可释放（否则 defer double-free）
				itemRef.Release()
				res.TotalFailed++
				continue
			}

			if cfg.CopyOnly {
				err = r.bridge.CopyItem(itemRef.Dispatch(), targetFolder)
			} else {
				err = r.bridge.MoveItem(itemRef.Dispatch(), targetFolder)
			}
			if err != nil {
				r.logger.Error("移动/复制失败", zap.Error(err))
				res.TotalFailed++
			} else {
				if forceMigrate {
					res.TotalMigrated++
					r.logger.Info("Migrated mail", zap.String("subject", m.subject), zap.String("source_pst", currentPSTName), zap.String("target_pst", targetPSTName))
				} else {
					res.TotalRectified++
					r.logger.Info("Rectified mail", zap.String("subject", m.subject), zap.String("source_pst", currentPSTName), zap.String("target_pst", targetPSTName))
				}
			}

			comutil.SafeRelease(targetFolder)
			// targetPSTRoot 在 defer 中统一释放
			itemRef.Release()

			if progressCb != nil {
				progressCb(2, res.TotalRectified+res.TotalMigrated, res.TotalRectified, currentPSTName)
			}

			time.Sleep(time.Duration(cfg.MoveIntervalMs) * time.Millisecond)
		}
	}
}
