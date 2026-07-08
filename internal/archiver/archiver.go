package archiver

import (
	"context"
	"fmt"
	"os"
	"sync"
	"time"

	"github.com/go-ole/go-ole"
	"go.uber.org/zap"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/monitor"
	"outlook-archiver/internal/outlook"
	"outlook-archiver/internal/router"
	"outlook-archiver/pkg/comutil"
)

type MailError struct {
	Subject string
	Err     error
}

// ArchiveResult 归档执行结果
type ArchiveResult struct {
	TotalMatched int
	TotalMoved   int
	TotalFailed  int
	TotalSkipped int
	Duration     time.Duration
	Errors       []MailError
}

type ArchiveOptions struct {
	MaxBatchSize int           // 0 = 不限制（全量整理模式）
	SafeDelay    time.Duration // 0 = 不延迟（全量整理模式）
	RetainDays   int           // 0 = 不保留最近的邮件，全部归档
	MoveInterval time.Duration
	DryRun       bool // true = 仅日志，不执行 Move
	CopyOnly     bool // true = 仅复制，不执行 Move
}

type Archiver struct {
	cfgProvider func() config.Config // 实时获取当前生效配置（支持热重载）
	bridge      *outlook.COMBridge
	logger      *zap.Logger

	pendingMu      sync.Mutex
	pendingDeletes []string // 待删除的 PST 文件路径
}

// getCfg 返回当前生效配置的最新副本（非启动时快照），保证热重载后归档逻辑使用新配置
func (a *Archiver) getCfg() config.Config {
	if a.cfgProvider == nil {
		return config.Config{}
	}
	return a.cfgProvider()
}

// checkPSTSize 检查目标 PST 文件大小，接近/超过阈值时输出告警日志（需求 §5.3）
func (a *Archiver) checkPSTSize(quarter router.QuarterInfo) {
	path := quarter.PSTFilePath(a.getCfg().PSTRootPath)
	status, err := monitor.CheckPSTSize(path)
	if err != nil {
		a.logger.Warn("检查 PST 大小失败", zap.String("path", path), zap.Error(err))
		return
	}
	switch status {
	case monitor.PSTSizeWarning:
		a.logger.Warn("单 PST 文件接近 20GB 上限，建议尽快拆分或迁移", zap.String("path", path), zap.Int64("size_gb_threshold", 15))
	case monitor.PSTSizeCritical:
		a.logger.Error("单 PST 文件超过 20GB，存在损坏风险，请立即拆分或迁移", zap.String("path", path), zap.Int64("size_gb_threshold", 20))
	}
}

func NewArchiver(cfgProvider func() config.Config, bridge *outlook.COMBridge, logger *zap.Logger) *Archiver {
	return &Archiver{
		cfgProvider: cfgProvider,
		bridge:      bridge,
		logger:      logger,
	}
}

// AddPendingDelete 将 PST 文件路径加入待删除列表
func (a *Archiver) AddPendingDelete(path string) {
	a.pendingMu.Lock()
	defer a.pendingMu.Unlock()
	a.pendingDeletes = append(a.pendingDeletes, path)
	a.logger.Info("已加入待删除列表", zap.String("pst", path))
}

// GetAndClearPendingDeletes 获取所有待删除的 PST 文件列表，并清空原列表
func (a *Archiver) GetAndClearPendingDeletes() []string {
	a.pendingMu.Lock()
	defer a.pendingMu.Unlock()
	if len(a.pendingDeletes) == 0 {
		return nil
	}
	paths := make([]string, len(a.pendingDeletes))
	copy(paths, a.pendingDeletes)
	a.pendingDeletes = nil
	return paths
}

// ExecuteIndependentDeletion 执行独立的空文件删除流程：关 Outlook、删文件、拉起 Outlook
func (a *Archiver) ExecuteIndependentDeletion() {
	pending := a.GetAndClearPendingDeletes()
	if len(pending) == 0 {
		return
	}
	a.logger.Info("开始执行独立删除流程，准备关闭 Outlook", zap.Int("count", len(pending)))

	// 0. 即将重启 Outlook，预先令 Namespace 缓存失效，避免悬空代理
	//    （QuitOutlook 内部也会失效，此处显式调用作为防御性兜底）
	a.bridge.InvalidateNamespaceCache()

	// 1. 尝试优雅关闭
	a.bridge.QuitOutlook()

	// 等待并检测是否真的关闭
	closed := false
	for i := 0; i < 5; i++ {
		time.Sleep(2 * time.Second)
		if !outlook.IsOutlookRunning() {
			closed = true
			break
		}
	}

	// 2. 强制杀进程
	if !closed {
		a.logger.Warn("优雅退出 Outlook 失败，执行强杀")
		outlook.ForceKillOutlook()
		time.Sleep(2 * time.Second)
	}

	// 3. 删除文件
	for _, p := range pending {
		err := os.Remove(p) // 注意需要确保 os 包被引入
		if err != nil {
			a.logger.Error("物理删除 PST 失败", zap.String("pst", p), zap.Error(err))
		} else {
			a.logger.Info("物理删除 PST 成功", zap.String("pst", p))
		}
	}

	// 4. 重启 Outlook
	a.logger.Info("准备重新启动 Outlook")
	outlook.StartOutlook()
}

// GetMountedPSTs delegates to COMBridge to fetch mounted PSTs
func (a *Archiver) GetMountedPSTs() ([]string, error) {
	return a.bridge.GetMountedPSTs()
}

// BuildRestrictFilter 根据文件夹类型构造 DASL 过滤条件
func BuildRestrictFilter(timeField string, cutoffTime time.Time) string {
	// 格式化为 Outlook 可接受的时间字符串
	timeStr := cutoffTime.Format("2006-01-02 15:04:05")
	return fmt.Sprintf("[%s] < '%s'", timeField, timeStr)
}

func CalcCutoffTime(safeDelay time.Duration, retainDays int) time.Time {
	now := time.Now()
	cutoff := now.Add(-safeDelay)
	if retainDays > 0 {
		retainCutoff := now.AddDate(0, 0, -retainDays)
		if retainCutoff.Before(cutoff) {
			cutoff = retainCutoff
		}
	}
	return cutoff
}

// Archive 执行一次常规归档
func (a *Archiver) Archive(ctx context.Context, opts ArchiveOptions) (*ArchiveResult, error) {
	start := time.Now()
	cfg := a.getCfg()
	res := &ArchiveResult{
		Errors: make([]MailError, 0),
	}

	if !outlook.IsOutlookRunning() {
		a.logger.Info("Outlook is not running, skipping archive")
		res.Duration = time.Since(start)
		return res, nil
	}

	a.logger.Info("正在扫描整个邮箱文件夹树，请稍候...")
	folders, err := a.bridge.WalkMailboxFolders(&cfg)
	if err != nil {
		a.logger.Error("Failed to walk mailbox folders", zap.Error(err))
		return nil, err
	}
	defer func() {
		for _, f := range folders {
			if f.Dispatch != nil {
				comutil.SafeRelease(f.Dispatch)
			}
		}
	}()

	for _, folder := range folders {
		if ctx.Err() != nil {
			return res, ctx.Err()
		}

		var moved, failed int
		_ = a.bridge.SubmitWithContext(ctx, func() error {
			moved, failed = a.processFolder(ctx, folder, opts, res)
			return nil
		})
		res.TotalMoved += moved
		res.TotalFailed += failed

		if opts.MaxBatchSize > 0 && res.TotalMoved >= opts.MaxBatchSize {
			a.logger.Info("Reached max batch size, stopping archive")
			break
		}
	}
	res.Duration = time.Since(start)
	return res, nil
}

// compactMeta 分块流式遍历的紧凑快照项，仅保留处理所需的最小信息。
// time 用 int64 时间戳压缩；分块方案下内存峰值由 blockSize 决定，
// 与文件夹邮件总数解耦（原 mailMeta 全量快照已移除）。
type compactMeta struct {
	entryID string
	subject string
	ts      int64
}

func (a *Archiver) processFolder(ctx context.Context, folder outlook.FolderInfo, opts ArchiveOptions, res *ArchiveResult) (int, int) {
	moved := 0
	failed := 0
	cfg := a.getCfg()

	a.logger.Info("正在获取文件夹内的邮件集合并应用过滤条件", zap.String("folder", folder.FullPath))
	restricted, count, cleanup, err := a.getRestrictedItems(folder, opts)
	if err != nil {
		a.logger.Error("获取过滤后邮件失败", zap.String("folder", folder.FullPath), zap.Error(err))
		return moved, failed
	}
	defer cleanup()
	res.TotalMatched += count

	if count == 0 {
		return moved, failed
	}

	a.logger.Info("正在请求 Outlook 对邮件进行排序，该操作可能比较耗时", zap.String("folder", folder.FullPath))
	// M5：显式排序，保证倒序遍历语义正确（Items 默认顺序不保证按时间）
	if err := a.bridge.SortItems(restricted, folder.TimeField, true); err != nil {
		a.logger.Warn("排序集合失败，回退为索引顺序", zap.String("folder", folder.FullPath), zap.Error(err))
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

	// M5：从末尾向头部分块倒序处理，内存峰值由 blockSize 决定（默认约 0.4MB），
	// 与文件夹邮件总数解耦；MaxBatchSize 可在分块循环开头提前终止，无需全量扫描。
	for end := count; end >= 1; end -= blockSize {
		if ctx.Err() != nil {
			res.TotalSkipped += end
			break
		}
		if opts.MaxBatchSize > 0 && res.TotalMoved+moved >= opts.MaxBatchSize {
			res.TotalSkipped += end
			break
		}

		start := end - blockSize + 1
		if start < 1 {
			start = 1
		}

		a.logger.Info(fmt.Sprintf("正在拉取邮件数据快照 [%d - %d] (共 %d 封)，请耐心等待拉取完成...", start, end, end-start+1), zap.String("folder", folder.FullPath))
		// 收集本块紧凑快照（正序索引遍历，仅 entryID + subject + 时间戳）
		block := make([]compactMeta, 0, end-start+1)
		for i := start; i <= end; i++ {
			if ctx.Err() != nil {
				break
			}
			item, err := a.bridge.GetItem(restricted, i)
			if err != nil || item == nil {
				continue
			}
			entryID, errID := a.bridge.GetEntryID(item)
			if errID != nil || entryID == "" {
				a.logger.Warn("Failed to get EntryID in snapshot", zap.Error(errID))
				comutil.SafeRelease(item)
				continue
			}
			subject := a.bridge.GetSubject(item)
			mailTime, errTime := a.bridge.GetMailTime(item, folder.TimeField)
			if errTime != nil {
				a.logger.Warn("Failed to get mail time in snapshot", zap.String("subject", subject), zap.Error(errTime))
				comutil.SafeRelease(item)
				continue
			}
			comutil.SafeRelease(item)
			
			a.logger.Debug("成功读取单封邮件快照", zap.Int("index", i), zap.String("subject", subject), zap.Time("mail_time", mailTime))
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
				res.TotalSkipped += (j + 1)
				break
			}
			if opts.MaxBatchSize > 0 && res.TotalMoved+moved >= opts.MaxBatchSize {
				res.TotalSkipped += (j + 1)
				break
			}

			mailItem, err := a.bridge.GetItemFromID(m.entryID)
			if err != nil || mailItem == nil {
				a.logger.Warn("Failed to get item by EntryID", zap.String("subject", m.subject), zap.Error(err))
				failed++
				continue
			}
			itemRef := comutil.NewCOMRef(mailItem, "mail_"+m.entryID)

			mailTime := time.Unix(m.ts, 0)
			quarter := router.CalcQuarter(mailTime)

			if opts.DryRun {
				action := "move"
				if opts.CopyOnly {
					action = "copy"
				}
				a.logger.Info("[DRY RUN] Would "+action+" mail",
					zap.String("subject", m.subject),
					zap.Time("mail_time", mailTime),
					zap.String("source_folder", folder.FullPath),
					zap.String("target_pst", quarter.PSTFileName()),
				)
				moved++
				itemRef.Release()
				time.Sleep(opts.MoveInterval)
				continue
			}

			// Ensure PST is mounted and Folder exists
			pstRoot, ok := pstCache[quarter.PSTFileName()]
			if !ok {
				pstRoot, err = a.bridge.EnsurePSTMounted(quarter, cfg.PSTRootPath)
				if err != nil {
					a.logger.Error("Failed to ensure PST mounted", zap.Error(err))
					res.Errors = append(res.Errors, MailError{Subject: m.subject, Err: err})
					itemRef.Release()
					failed++
					continue
				}
				pstCache[quarter.PSTFileName()] = pstRoot
				a.checkPSTSize(quarter) // M12：挂载目标 PST 后检查大小并告警
			}

			folderPath := folder.FullPath
			targetFolder, err := a.bridge.EnsurePSTFolder(pstRoot, folderPath)
			if err != nil {
				a.logger.Error("Failed to ensure PST folder", zap.Error(err))
				res.Errors = append(res.Errors, MailError{Subject: m.subject, Err: err})
				itemRef.Release()
				failed++
				continue
			}

			if opts.CopyOnly {
				err = a.bridge.CopyItem(itemRef.Dispatch(), targetFolder)
				if err != nil {
					a.logger.Error("Failed to copy item", zap.String("subject", m.subject), zap.Error(err))
					res.Errors = append(res.Errors, MailError{Subject: m.subject, Err: err})
					failed++
				} else {
					a.logger.Info("Copied mail", zap.String("subject", m.subject), zap.String("target_pst", quarter.PSTFileName()))
					moved++
				}
			} else {
				err = a.bridge.MoveItem(itemRef.Dispatch(), targetFolder)
				if err != nil {
					a.logger.Error("Failed to move item", zap.String("subject", m.subject), zap.Error(err))
					res.Errors = append(res.Errors, MailError{Subject: m.subject, Err: err})
					failed++
				} else {
					a.logger.Info("Moved mail", zap.String("subject", m.subject), zap.String("target_pst", quarter.PSTFileName()))
					moved++
				}
			}

			comutil.SafeRelease(targetFolder)
			// 不在这里 Release pstRoot，在 defer 中统一释放
			itemRef.Release()

			time.Sleep(opts.MoveInterval)
		}
	}

	return moved, failed
}

func (a *Archiver) getRestrictedItems(folder outlook.FolderInfo, opts ArchiveOptions) (*ole.IDispatch, int, func(), error) {
	itemsVar, err := comutil.SafeGetProperty(folder.Dispatch, "Items")
	if err != nil {
		return nil, 0, nil, err
	}
	items := itemsVar.ToIDispatch()
	if items == nil {
		itemsVar.Clear()
		return nil, 0, nil, fmt.Errorf("Items is nil")
	}

	var filter string
	if opts.SafeDelay > 0 || opts.RetainDays > 0 {
		cutoffTime := CalcCutoffTime(opts.SafeDelay, opts.RetainDays)
		filter = BuildRestrictFilter(folder.TimeField, cutoffTime)
	}

	var restricted *ole.IDispatch
	var restrictedVar *ole.VARIANT
	if filter != "" {
		restrictedVar, err = comutil.SafeCallMethod(items, "Restrict", filter)
		if err != nil {
			comutil.SafeRelease(items)
			itemsVar.Clear()
			return nil, 0, nil, err
		}
		restricted = restrictedVar.ToIDispatch()
	} else {
		items.AddRef()
		restricted = items
	}

	if restricted == nil {
		if restrictedVar != nil {
			restrictedVar.Clear()
		}
		comutil.SafeRelease(items)
		itemsVar.Clear()
		return nil, 0, nil, fmt.Errorf("Restricted items is nil")
	}

	countVar, err := comutil.SafeGetProperty(restricted, "Count")
	if err != nil {
		comutil.SafeRelease(restricted)
		if restrictedVar != nil {
			restrictedVar.Clear()
		}
		comutil.SafeRelease(items)
		itemsVar.Clear()
		return nil, 0, nil, err
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

	cleanup := func() {
		countVar.Clear()
		comutil.SafeRelease(restricted)
		if restrictedVar != nil {
			restrictedVar.Clear()
		}
		comutil.SafeRelease(items)
		itemsVar.Clear()
	}

	return restricted, count, cleanup, nil
}
