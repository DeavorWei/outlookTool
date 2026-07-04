package outlook

import (
	"strings"

	"github.com/go-ole/go-ole"
	"outlook-archiver/internal/config"
	"outlook-archiver/pkg/comutil"
)

type FolderType int

const (
	FolderTypeDefault FolderType = iota
	FolderTypeCustom
)

type FolderInfo struct {
	Name       string
	FullPath   string
	FolderType FolderType
	TimeField  string
	Dispatch   *ole.IDispatch
}

// WalkMailboxFolders 遍历主邮箱的所有文件夹
func (b *COMBridge) WalkMailboxFolders(cfg *config.Config) ([]FolderInfo, error) {
	var results []FolderInfo
	err := b.Submit(func() error {
		ns, err := b.getNamespace()
		if err != nil {
			return err
		}
		defer comutil.SafeRelease(ns)

		var sentItemsID string
		sentFolderVar, err := comutil.SafeCallMethod(ns, "GetDefaultFolder", 5) // olFolderSentMail
		if err == nil && sentFolderVar.Value() != nil {
			sentFolder := sentFolderVar.ToIDispatch()
			idVar, err := comutil.SafeGetProperty(sentFolder, "EntryID")
			if err == nil && idVar.Value() != nil {
				sentItemsID = idVar.ToString()
				idVar.Clear()
			}
			sentFolderVar.Clear()
		}

		// 获取默认邮箱的 Inbox
		inboxVar, err := comutil.SafeCallMethod(ns, "GetDefaultFolder", 6) // olFolderInbox = 6
		if err != nil {
			return err
		}
		defer inboxVar.Clear()
		inbox := inboxVar.ToIDispatch()

		parentVar, err := comutil.SafeGetProperty(inbox, "Parent")
		if err != nil {
			return err
		}
		defer parentVar.Clear()
		rootFolder := parentVar.ToIDispatch()

		return b.walkFoldersRecursive(rootFolder, "", cfg, &results, true, sentItemsID)
	})
	return results, err
}

// WalkPSTFolders 遍历指定 PST 内的所有文件夹
func (b *COMBridge) WalkPSTFolders(pstRoot *ole.IDispatch) ([]FolderInfo, error) {
	var results []FolderInfo
	// PST 遍历时不应用配置文件中的 include/exclude，全量遍历
	err := b.Submit(func() error {
		return b.walkFoldersRecursive(pstRoot, "", nil, &results, false, "")
	})
	return results, err
}

func (b *COMBridge) walkFoldersRecursive(currentFolder *ole.IDispatch, currentPath string, cfg *config.Config, results *[]FolderInfo, checkSystem bool, sentItemsID string) error {
	foldersVar, err := comutil.SafeGetProperty(currentFolder, "Folders")
	if err != nil {
		return nil
	}
	defer foldersVar.Clear()
	folders := foldersVar.ToIDispatch()

	countVar, err := comutil.SafeGetProperty(folders, "Count")
	if err != nil {
		return nil
	}
	count := int(countVar.Val)
	countVar.Clear()

	for i := 1; i <= count; i++ {
		folderVar, err := comutil.SafeCallMethod(folders, "Item", i)
		if err != nil {
			continue
		}
		folder := folderVar.ToIDispatch()
		folder.AddRef()
		folderVar.Clear()
		
		nameVar, err := comutil.SafeGetProperty(folder, "Name")
		if err != nil {
			comutil.SafeRelease(folder)
			continue
		}
		name := nameVar.ToString()
		nameVar.Clear() // Fix memory leak for BSTR

		fullPath := name
		if currentPath != "" {
			fullPath = currentPath + "/" + name
		}

		if checkSystem {
			if b.isSystemReserved(folder, name, fullPath) {
				comutil.SafeRelease(folder)
				continue
			}
		}

		if cfg != nil {
			if shouldExclude(fullPath, cfg) {
				comutil.SafeRelease(folder)
				continue
			}
			if cfg.ArchiveMode == "list" {
				if isPathIncluded(fullPath, cfg) {
					folder.AddRef()
					*results = append(*results, FolderInfo{
						Name:       name,
						FullPath:   fullPath,
						FolderType: FolderTypeCustom,
						TimeField:  getTimeField(folder, name, sentItemsID),
						Dispatch:   folder,
					})
				} else if isPathPrefixOfIncluded(fullPath, cfg) {
					// 仅进入子文件夹，不将自己加入 results
				} else {
					comutil.SafeRelease(folder)
					continue
				}
			} else {
				folder.AddRef()
				*results = append(*results, FolderInfo{
					Name:       name,
					FullPath:   fullPath,
					FolderType: FolderTypeCustom,
					TimeField:  getTimeField(folder, name, sentItemsID),
					Dispatch:   folder,
				})
			}
		} else {
			folder.AddRef()
			*results = append(*results, FolderInfo{
				Name:       name,
				FullPath:   fullPath,
				FolderType: FolderTypeCustom,
				TimeField:  getTimeField(folder, name, sentItemsID),
				Dispatch:   folder,
			})
		}

		_ = b.walkFoldersRecursive(folder, fullPath, cfg, results, checkSystem, sentItemsID)
		comutil.SafeRelease(folder)
	}

	return nil
}

func (b *COMBridge) isSystemReserved(folder *ole.IDispatch, name, fullPath string) bool {
	if strings.HasPrefix(fullPath, "同步问题") || strings.HasPrefix(fullPath, "Sync Issues") {
		return true
	}
	
	// DefaultItemType: 0 = MailItem. 排除非邮件文件夹
	itemTypeVar, err := comutil.SafeGetProperty(folder, "DefaultItemType")
	if err == nil {
		isMail := int(itemTypeVar.Val) == 0
		itemTypeVar.Clear()
		if !isMail {
			return true
		}
	}
	
	lowerName := strings.ToLower(name)
	// 根据要求排除的内置保留文件夹名称
	reservedNames := []string{
		"已删除邮件", "deleted items",
		"发件箱", "outbox",
		"草稿", "drafts",
		"垃圾邮件", "junk e-mail",
		"垃圾电邮",
	}
	for _, rn := range reservedNames {
		if lowerName == rn {
			return true
		}
	}
	
	return false
}

func shouldExclude(path string, cfg *config.Config) bool {
	for _, ex := range cfg.ExcludeFolders {
		if strings.EqualFold(path, ex) || strings.HasPrefix(strings.ToLower(path), strings.ToLower(ex)+"/") {
			return true
		}
	}
	return false
}

func isPathIncluded(path string, cfg *config.Config) bool {
	for _, inc := range cfg.IncludeFolders {
		if strings.EqualFold(path, inc) || strings.HasPrefix(strings.ToLower(path), strings.ToLower(inc)+"/") {
			return true
		}
	}
	return false
}

func isPathPrefixOfIncluded(path string, cfg *config.Config) bool {
	for _, inc := range cfg.IncludeFolders {
		if strings.HasPrefix(strings.ToLower(inc), strings.ToLower(path)+"/") {
			return true
		}
	}
	return false
}

func getTimeField(folder *ole.IDispatch, name string, sentItemsID string) string {
	if sentItemsID != "" {
		idVar, err := comutil.SafeGetProperty(folder, "EntryID")
		if err == nil && idVar.Value() != nil {
			folderID := idVar.ToString()
			idVar.Clear()
			if folderID == sentItemsID {
				return "SentOn"
			}
		}
	} else {
		lowerName := strings.ToLower(name)
		if lowerName == "已发送邮件" || lowerName == "sent items" {
			return "SentOn"
		}
	}
	return "ReceivedTime"
}

// GetDefaultMailboxRoot 获取主邮箱（OST）的根目录
func (b *COMBridge) GetDefaultMailboxRoot() (*ole.IDispatch, error) {
	var root *ole.IDispatch
	err := b.Submit(func() error {
		ns, err := b.getNamespace()
		if err != nil {
			return err
		}
		defer comutil.SafeRelease(ns)

		inboxVar, err := comutil.SafeCallMethod(ns, "GetDefaultFolder", 6) // olFolderInbox
		if err != nil {
			return err
		}
		defer inboxVar.Clear()
		inbox := inboxVar.ToIDispatch()

		parentVar, err := comutil.SafeGetProperty(inbox, "Parent")
		if err != nil {
			return err
		}
		root = parentVar.ToIDispatch()
		return nil
	})
	return root, err
}
