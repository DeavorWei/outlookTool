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
		_ = a.bridge.RemoveStore(pstRoot)
		comutil.SafeRelease(pstRoot)

		if deleteEmptyPST && isEmpty {
			a.logger.Info("PST 已清空，准备物理删除", zap.String("pst", pstPath))
			if a.cfg.DryRun {
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
	pathSet := make(map[string]bool)
	
	mounted, err := a.bridge.GetMountedPSTs()
	if err != nil {
		a.logger.Warn("获取挂载 PST 列表失败", zap.Error(err))
	} else {
		for _, p := range mounted {
			pathSet[strings.ToLower(p)] = true
		}
	}
	
	for _, dir := range a.cfg.LegacyPSTScanPaths {
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
				pathSet[strings.ToLower(fullPath)] = true
			}
		}
	}
	
	var list []string
	for p := range pathSet {
		list = append(list, p)
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
	return a.bridge.Submit(func() error {
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
				if a.cfg.DryRun {
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
		if a.cfg.DryRun {
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
		if a.cfg.MoveIntervalMs > 0 {
			time.Sleep(time.Duration(a.cfg.MoveIntervalMs) * time.Millisecond)
		}
	}
	
	return nil
	})
}

func (a *Archiver) buildDuplicateSet(ctx context.Context, folder *ole.IDispatch, timeField string) map[string]bool {
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

// isFileLocked 检查文件是否被其他进程占用（尝试以读写模式独占打开）
func isFileLocked(path string) bool {
	f, err := os.OpenFile(path, os.O_RDWR, 0666)
	if err != nil {
		return true
	}
	f.Close()
	return false
}

// deleteFileWithRetry 在独立协程中删除文件
// 先等待 30 秒让 Outlook 释放句柄，之后每隔 30 秒重试一次，共尝试 3 次，总等待时间约 2 分钟
func (a *Archiver) deleteFileWithRetry(path string) {
	a.logger.Info("启动异步删除，等待 Outlook 释放文件句柄...", zap.String("pst", path))
	time.Sleep(30 * time.Second)

	const (
		maxAttempts   = 3
		retryInterval = 30 * time.Second
	)

	for attempt := 1; attempt <= maxAttempts; attempt++ {
		// 在删除前验证文件是否仍被占用
		if isFileLocked(path) {
			a.logger.Warn("文件仍被占用，等待后重试",
				zap.String("pst", path),
				zap.Int("attempt", attempt),
				zap.Int("maxAttempts", maxAttempts),
				zap.Duration("wait", retryInterval))
			time.Sleep(retryInterval)
			continue
		}

		// 文件未被占用，尝试删除
		err := os.Remove(path)
		if err == nil {
			a.logger.Info("空 PST 删除成功", zap.String("pst", path))
			return
		}

		// 删除失败（可能在检查与删除之间被重新锁定）
		a.logger.Warn("删除文件失败，等待后重试",
			zap.String("pst", path),
			zap.Int("attempt", attempt),
			zap.Error(err),
			zap.Duration("wait", retryInterval))
		time.Sleep(retryInterval)
	}

	a.logger.Error("删除空 PST 最终失败（已耗尽所有重试）",
		zap.String("pst", path),
		zap.Int("totalAttempts", maxAttempts))
}
