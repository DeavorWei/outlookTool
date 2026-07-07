package archiver

import (
	"context"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"go.uber.org/zap"

	"outlook-archiver/internal/outlook"
	"outlook-archiver/pkg/comutil"
)

// Restore 提取挂载的 PST 及历史目录，将其中的邮件还原到默认 OST 主邮箱中。
func (a *Archiver) Restore(ctx context.Context, deleteEmptyPST, deleteDuplicates bool) (*ArchiveResult, error) {
	cfg := a.getCfg()
	a.logger.Info("开始准备还原...", zap.Bool("deleteEmpty", deleteEmptyPST), zap.Bool("deleteDup", deleteDuplicates))

	// 1. 获取要处理的 PST 列表
	pstPaths, err := a.collectPSTsToRestore()
	if err != nil {
		return nil, fmt.Errorf("收集 PST 失败: %w", err)
	}

	if len(pstPaths) == 0 {
		a.logger.Info("没有找到任何 PST 文件")
		return &ArchiveResult{}, nil
	}

	// 2. 获取 OST 根目录
	ostRoot, err := a.bridge.GetDefaultMailboxRoot()
	if err != nil {
		return nil, fmt.Errorf("获取 OST 根目录失败: %w", err)
	}
	defer comutil.SafeRelease(ostRoot)

	result := &ArchiveResult{}

	// 3. 逐个处理 PST
	for _, pstPath := range pstPaths {
		select {
		case <-ctx.Done():
			return result, ctx.Err()
		default:
		}

		a.logger.Info("开始处理 PST", zap.String("pst", pstPath))

		// 挂载 PST
		pstRoot, err := a.bridge.EnsurePSTMountedByPath(pstPath)
		if err != nil {
			a.logger.Error("挂载 PST 失败", zap.String("pst", pstPath), zap.Error(err))
			result.Errors = append(result.Errors, MailError{Subject: pstPath, Err: err})
			continue
		}

		isEmpty, err := a.restorePST(ctx, pstPath, pstRoot, ostRoot, deleteEmptyPST, deleteDuplicates, result)

		// 卸载
		if removeErr := a.bridge.RemoveStore(pstRoot); removeErr != nil {
			a.logger.Warn("卸载 PST 失败", zap.String("pst", pstPath), zap.Error(removeErr))
		}
		comutil.SafeRelease(pstRoot)

		if deleteEmptyPST && isEmpty {
			a.logger.Info("PST 已清空，准备物理删除", zap.String("pst", pstPath))
			if cfg.DryRun {
				a.logger.Info("[Dry Run] 模拟删除空 PST 文件", zap.String("pst", pstPath))
			} else {
				// 加入待删除列表，由定时任务完成后统一处理
				a.AddPendingDelete(pstPath)
			}
		}

		if err != nil {
			a.logger.Error("还原 PST 时出错", zap.String("pst", pstPath), zap.Error(err))
		}
	}

	return result, nil
}

// collectPSTsToRestore 收集当前挂载的 PST 以及遗留目录下的所有 PST
func (a *Archiver) collectPSTsToRestore() ([]string, error) {
	cfg := a.getCfg()
	pathSet := make(map[string]string)

	mounted, err := a.bridge.GetMountedPSTs()
	if err != nil {
		a.logger.Warn("获取挂载 PST 列表失败", zap.Error(err))
	} else {
		for _, p := range mounted {
			pathSet[strings.ToLower(p)] = p
		}
	}

	for _, dir := range cfg.LegacyPSTScanPaths {
		if dir == "" {
			continue
		}
		entries, err := os.ReadDir(dir)
		if err != nil {
			a.logger.Warn("读取历史目录失败", zap.String("dir", dir), zap.Error(err))
			continue
		}
		for _, entry := range entries {
			if !entry.IsDir() && strings.HasSuffix(strings.ToLower(entry.Name()), ".pst") {
				fullPath := fmt.Sprintf("%s\\%s", dir, entry.Name())
				pathSet[strings.ToLower(fullPath)] = fullPath
			}
		}
	}

	var list []string
	for _, originalPath := range pathSet {
		list = append(list, originalPath)
	}
	return list, nil
}

func (a *Archiver) restorePST(ctx context.Context, pstPath string, pstRoot, ostRoot *ole.IDispatch, deleteEmpty, deleteDup bool, res *ArchiveResult) (bool, error) {
	folders, err := a.bridge.WalkPSTFolders(pstRoot)
	if err != nil {
		return false, err
	}

	for _, f := range folders {
		select {
		case <-ctx.Done():
			return false, ctx.Err()
		default:
		}

		a.logger.Info("处理 PST 文件夹", zap.String("folder", f.FullPath))

		ostFolder, err := a.bridge.EnsurePSTFolder(ostRoot, f.FullPath)
		if err != nil {
			a.logger.Error("在 OST 中创建目标文件夹失败", zap.String("folder", f.FullPath), zap.Error(err))
			continue
		}

		err = a.restoreFolderItems(ctx, f, ostFolder, deleteDup, res)
		comutil.SafeRelease(ostFolder)
		if err != nil {
			a.logger.Error("处理文件夹邮件失败", zap.String("folder", f.FullPath), zap.Error(err))
		}
	}

	isEmpty := true
	if deleteEmpty {
		for _, f := range folders {
			count, err := a.bridge.GetFolderItemCount(f.Dispatch)
			if err != nil || count > 0 {
				isEmpty = false
				break
			}
		}
	}

	// 释放所有的文件夹引用，防止占用 PST 导致无法卸载或删除
	for _, f := range folders {
		if f.Dispatch != nil {
			comutil.SafeRelease(f.Dispatch)
		}
	}

	if deleteEmpty {
		if !isEmpty {
			a.logger.Info("PST 不为空，保留", zap.String("pst", pstPath))
		}
	}

	return isEmpty, nil
}

func (a *Archiver) restoreFolderItems(ctx context.Context, pstFolder outlook.FolderInfo, ostFolder *ole.IDispatch, deleteDup bool, res *ArchiveResult) error {
	cfg := a.getCfg()
	return a.bridge.SubmitWithContext(ctx, func() error {
		// 收集 OST 现有的 key 用于去重
		var ostSet map[string]bool
		if deleteDup {
			ostSet = a.buildDuplicateSet(ctx, ostFolder, pstFolder.TimeField)
		}

		itemsVar, err := comutil.SafeGetProperty(pstFolder.Dispatch, "Items")
		if err != nil || itemsVar.Value() == nil {
			return err
		}
		defer itemsVar.Clear()
		items := itemsVar.ToIDispatch()
		defer comutil.SafeRelease(items)

		countVar, err := comutil.SafeGetProperty(items, "Count")
		if err != nil {
			return err
		}
		count := int(countVar.Val)
		countVar.Clear()

		for i := count; i >= 1; i-- { // 倒序处理防止删/移动导致下标错乱
			select {
			case <-ctx.Done():
				return ctx.Err()
			default:
			}

			itemVar, err := comutil.SafeCallMethod(items, "Item", i)
			if err != nil || itemVar.Value() == nil {
				continue
			}
			item := itemVar.ToIDispatch()

			res.TotalMatched++

			if deleteDup && ostSet != nil {
				subj := a.bridge.GetSubject(item)
				timeVal, _ := a.bridge.GetMailTime(item, pstFolder.TimeField)
				key := fmt.Sprintf("%s_%d", subj, timeVal.Unix())

				if ostSet[key] {
					// 重复了，删除 PST 中的
					a.logger.Debug("发现重复邮件，拟删除源", zap.String("subject", subj))
					if cfg.DryRun {
						a.logger.Info("[Dry Run] 模拟删除重复邮件", zap.String("subject", subj))
						res.TotalMoved++
					} else {
						if err := a.bridge.DeleteItem(item); err != nil {
							a.logger.Warn("删除重复邮件失败", zap.Error(err), zap.String("subject", subj))
						} else {
							a.logger.Info("成功删除重复邮件", zap.String("subject", subj))
							res.TotalMoved++ // 把处理掉的重复件也视为处理成功
						}
					}
					itemVar.Clear()
					continue
				} else {
					// 未重复，加入 set 以防止 PST 内部有多份
					ostSet[key] = true
				}
			}

			// 移动邮件
			subj := a.bridge.GetSubject(item)
			if cfg.DryRun {
				a.logger.Info("[Dry Run] 模拟移动邮件", zap.String("subject", subj))
				res.TotalMoved++
			} else {
				if err := a.bridge.MoveItem(item, ostFolder); err != nil {
					a.logger.Warn("移动邮件失败", zap.Error(err), zap.String("subject", subj))
					res.TotalFailed++
					res.Errors = append(res.Errors, MailError{Subject: subj, Err: err})
				} else {
					a.logger.Info("成功移动邮件", zap.String("subject", subj))
					res.TotalMoved++
				}
			}

			itemVar.Clear()
			if cfg.MoveIntervalMs > 0 {
				time.Sleep(time.Duration(cfg.MoveIntervalMs) * time.Millisecond)
			}
		}

		return nil
	})
}

func (a *Archiver) buildDuplicateSet(ctx context.Context, folder *ole.IDispatch, timeField string) map[string]bool {
	// 优先使用快速方式
	if fastSet, err := a.buildDuplicateSetFast(ctx, folder, timeField); err == nil && fastSet != nil {
		return fastSet
	}

	a.logger.Warn("使用 GetTable 构建索引失败，回退到慢速方法")

	set := make(map[string]bool)
	itemsVar, err := comutil.SafeGetProperty(folder, "Items")
	if err != nil || itemsVar.Value() == nil {
		return set
	}
	defer itemsVar.Clear()
	items := itemsVar.ToIDispatch()
	defer comutil.SafeRelease(items)

	countVar, err := comutil.SafeGetProperty(items, "Count")
	if err != nil {
		return set
	}
	count := int(countVar.Val)
	countVar.Clear()

	if count > 5000 {
		a.logger.Info("目标文件夹较大，正在全量构建去重索引，这可能需要一点时间...", zap.Int("total", count))
	}

	for i := 1; i <= count; i++ {
		select {
		case <-ctx.Done():
			return set
		default:
		}

		itemVar, err := comutil.SafeCallMethod(items, "Item", i)
		if err != nil || itemVar.Value() == nil {
			continue
		}
		item := itemVar.ToIDispatch()
		subj := a.bridge.GetSubject(item)
		timeVal, _ := a.bridge.GetMailTime(item, timeField)
		key := fmt.Sprintf("%s_%d", subj, timeVal.Unix())
		set[key] = true

		itemVar.Clear()

		if i%5000 == 0 {
			a.logger.Info("去重索引构建进度", zap.Int("processed", i), zap.Int("total", count))
		}
	}
	return set
}

func (a *Archiver) buildDuplicateSetFast(ctx context.Context, folder *ole.IDispatch, timeField string) (map[string]bool, error) {
	set := make(map[string]bool)

	tableVar, err := comutil.SafeCallMethod(folder, "GetTable")
	if err != nil || tableVar.Value() == nil {
		return nil, err
	}
	defer tableVar.Clear()
	table := tableVar.ToIDispatch()
	defer comutil.SafeRelease(table)

	colsVar, err := comutil.SafeGetProperty(table, "Columns")
	if err != nil || colsVar.Value() == nil {
		return nil, err
	}
	defer colsVar.Clear()
	cols := colsVar.ToIDispatch()
	defer comutil.SafeRelease(cols)

	comutil.SafeCallMethod(cols, "Add", "Subject")
	comutil.SafeCallMethod(cols, "Add", timeField)

	for {
		select {
		case <-ctx.Done():
			return set, ctx.Err()
		default:
		}

		endVar, err := comutil.SafeGetProperty(table, "EndOfTable")
		if err == nil && endVar.Value() != nil && endVar.Value().(bool) {
			endVar.Clear()
			break
		}
		if endVar != nil {
			endVar.Clear()
		}

		rowVar, err := comutil.SafeCallMethod(table, "GetNextRow")
		if err != nil || rowVar.Value() == nil {
			break
		}
		row := rowVar.ToIDispatch()

		var subj string
		subjItem, _ := comutil.SafeCallMethod(row, "Item", "Subject")
		if subjItem != nil && subjItem.Value() != nil {
			if s, ok := subjItem.Value().(string); ok {
				subj = s
			}
			subjItem.Clear()
		}

		var t time.Time
		timeItem, _ := comutil.SafeCallMethod(row, "Item", timeField)
		if timeItem != nil && timeItem.Value() != nil {
			t, _ = outlook.ParseTime(timeItem.Value())
			timeItem.Clear()
		}

		key := fmt.Sprintf("%s_%d", subj, t.Unix())
		set[key] = true

		comutil.SafeRelease(row)
		rowVar.Clear()
	}

	a.logger.Info("使用快速查询构建了目标索引", zap.Int("count", len(set)))
	return set, nil
}
