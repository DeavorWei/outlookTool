package archiver

import (
	"context"
	"fmt"
	"time"

	"go.uber.org/zap"
	"outlook-archiver/internal/config"
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
	MoveInterval time.Duration
	DryRun       bool          // true = 仅日志，不执行 Move
}

type Archiver struct {
	cfg    *config.Config
	bridge *outlook.COMBridge
	logger *zap.Logger
}

func NewArchiver(cfg *config.Config, bridge *outlook.COMBridge, logger *zap.Logger) *Archiver {
	return &Archiver{
		cfg:    cfg,
		bridge: bridge,
		logger: logger,
	}
}

// BuildRestrictFilter 根据文件夹类型构造 DASL 过滤条件
func BuildRestrictFilter(timeField string, cutoffTime time.Time) string {
	// 格式化为 Outlook 可接受的时间字符串
	timeStr := cutoffTime.Format("2006-01-02 15:04:05")
	return fmt.Sprintf("[%s] < '%s' AND [MessageClass] = 'IPM.Note'", timeField, timeStr)
}

func calcCutoffTime(safeDelay time.Duration) time.Time {
	return time.Now().Add(-safeDelay)
}

// Archive 执行一次常规归档
func (a *Archiver) Archive(ctx context.Context, opts ArchiveOptions) (*ArchiveResult, error) {
	start := time.Now()
	res := &ArchiveResult{
		Errors: make([]MailError, 0),
	}

	if !outlook.IsOutlookRunning() {
		a.logger.Info("Outlook is not running, skipping archive")
		res.Duration = time.Since(start)
		return res, nil
	}

	folders, err := a.bridge.WalkMailboxFolders(a.cfg)
	if err != nil {
		a.logger.Error("Failed to walk mailbox folders", zap.Error(err))
		return nil, err
	}

	for idx, folder := range folders {
		if ctx.Err() != nil {
			return res, ctx.Err()
		}

		var moved, failed int
		_ = a.bridge.Submit(func() error {
			moved, failed = a.processFolder(ctx, folder, opts, res)
			return nil
		})
		res.TotalMoved += moved
		res.TotalFailed += failed

		if opts.MaxBatchSize > 0 && res.TotalMoved >= opts.MaxBatchSize {
			a.logger.Info("Reached max batch size, stopping archive")
			for j := idx + 1; j < len(folders); j++ {
				var count int
				_ = a.bridge.Submit(func() error {
					count = a.countMatchedItems(folders[j], opts)
					return nil
				})
				res.TotalMatched += count
				res.TotalSkipped += count
			}
			break
		}
	}

	res.Duration = time.Since(start)
	return res, nil
}

func (a *Archiver) processFolder(ctx context.Context, folder outlook.FolderInfo, opts ArchiveOptions, res *ArchiveResult) (int, int) {
	moved := 0
	failed := 0

	itemsVar, err := comutil.SafeGetProperty(folder.Dispatch, "Items")
	if err != nil {
		a.logger.Error("Failed to get Items", zap.String("folder", folder.FullPath), zap.Error(err))
		return moved, failed
	}
	items := itemsVar.ToIDispatch()
	if items == nil {
		a.logger.Error("Items is nil", zap.String("folder", folder.FullPath))
		return moved, failed
	}
	defer comutil.SafeRelease(items)

	var filter string
	if opts.SafeDelay == 0 {
		filter = "[MessageClass] = 'IPM.Note'"
	} else {
		cutoffTime := calcCutoffTime(opts.SafeDelay)
		filter = BuildRestrictFilter(folder.TimeField, cutoffTime)
	}

	a.logger.Debug("Restrict filter", zap.String("folder", folder.FullPath), zap.String("filter", filter))

	restrictedVar, err := comutil.SafeCallMethod(items, "Restrict", filter)
	if err != nil {
		a.logger.Error("Failed to restrict Items", zap.String("folder", folder.FullPath), zap.Error(err))
		return moved, failed
	}
	restricted := restrictedVar.ToIDispatch()
	if restricted == nil {
		a.logger.Error("Restricted Items is nil", zap.String("folder", folder.FullPath))
		return moved, failed
	}
	defer comutil.SafeRelease(restricted)

	countVar, err := comutil.SafeGetProperty(restricted, "Count")
	if err != nil {
		a.logger.Error("Failed to get Count", zap.String("folder", folder.FullPath), zap.Error(err))
		return moved, failed
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
	res.TotalMatched += count

	for i := count; i >= 1; i-- {
		if ctx.Err() != nil {
			res.TotalSkipped += i
			break
		}
		if opts.MaxBatchSize > 0 && res.TotalMoved+moved >= opts.MaxBatchSize {
			res.TotalSkipped += i
			break
		}

		itemVar, err := comutil.SafeCallMethod(restricted, "Item", i)
		if err != nil {
			a.logger.Warn("Failed to get item", zap.Int("index", i), zap.Error(err))
			failed++
			continue
		}
		item := itemVar.ToIDispatch()
		if item == nil {
			a.logger.Warn("Item is nil", zap.Int("index", i))
			failed++
			continue
		}
		itemRef := comutil.NewCOMRef(item, fmt.Sprintf("mail_%d", i))

		// Read properties
		subject := a.bridge.GetSubject(itemRef.Dispatch())
		mailTime, err := a.bridge.GetMailTime(itemRef.Dispatch(), folder.TimeField)
		if err != nil {
			a.logger.Warn("Failed to get mail time", zap.String("subject", subject), zap.Error(err))
			itemRef.Release()
			failed++
			continue
		}

		quarter := router.CalcQuarter(mailTime)

		if opts.DryRun {
			a.logger.Info("[DRY RUN] Would move mail",
				zap.String("subject", subject),
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
		pstRoot, err := a.bridge.EnsurePSTMounted(quarter, a.cfg.PSTRootPath)
		if err != nil {
			a.logger.Error("Failed to ensure PST mounted", zap.Error(err))
			res.Errors = append(res.Errors, MailError{Subject: subject, Err: err})
			itemRef.Release()
			failed++
			continue
		}

		targetFolder, err := a.bridge.EnsurePSTFolder(pstRoot, folder.FullPath)
		if err != nil {
			a.logger.Error("Failed to ensure PST folder", zap.Error(err))
			res.Errors = append(res.Errors, MailError{Subject: subject, Err: err})
			comutil.SafeRelease(pstRoot)
			itemRef.Release()
			failed++
			continue
		}

		err = a.bridge.MoveItem(itemRef.Dispatch(), targetFolder)
		if err != nil {
			a.logger.Error("Failed to move item", zap.String("subject", subject), zap.Error(err))
			res.Errors = append(res.Errors, MailError{Subject: subject, Err: err})
			failed++
		} else {
			moved++
		}

		comutil.SafeRelease(targetFolder)
		comutil.SafeRelease(pstRoot)
		itemRef.Release()

		time.Sleep(opts.MoveInterval)
	}

	return moved, failed
}

func (a *Archiver) countMatchedItems(folder outlook.FolderInfo, opts ArchiveOptions) int {
	itemsVar, err := comutil.SafeGetProperty(folder.Dispatch, "Items")
	if err != nil {
		return 0
	}
	items := itemsVar.ToIDispatch()
	if items == nil {
		return 0
	}
	defer comutil.SafeRelease(items)

	var filter string
	if opts.SafeDelay == 0 {
		filter = "[MessageClass] = 'IPM.Note'"
	} else {
		cutoffTime := calcCutoffTime(opts.SafeDelay)
		filter = BuildRestrictFilter(folder.TimeField, cutoffTime)
	}

	restrictedVar, err := comutil.SafeCallMethod(items, "Restrict", filter)
	if err != nil {
		return 0
	}
	restricted := restrictedVar.ToIDispatch()
	if restricted == nil {
		return 0
	}
	defer comutil.SafeRelease(restricted)

	countVar, err := comutil.SafeGetProperty(restricted, "Count")
	if err != nil {
		return 0
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
	return count
}
