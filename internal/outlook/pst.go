package outlook

import (
	"fmt"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"outlook-archiver/internal/router"
	"outlook-archiver/pkg/comutil"
)

// getNamespace 获取 MAPI Namespace（单例缓存）。
//
// 缓存策略：
//   - 命中：对 cachedNS 执行 AddRef 后返回，调用方仍需 defer SafeRelease(ns)。
//   - 未命中：获取 Application 与 Namespace 并缓存，缓存自身持有一份引用。
//
// 线程约束：仅可在 COM 线程调用。所有调用方均已位于 b.Submit 闭包内，
// 由 SubmitWithContext 的线程 ID 检查保证重入时内联执行，不会死锁。
func (b *COMBridge) getNamespace() (*ole.IDispatch, error) {
	// 1. 缓存命中：存活校验通过则 AddRef 返回
	if b.cachedNS != nil {
		if IsOutlookRunning() {
			b.cachedNS.AddRef()
			return b.cachedNS, nil
		}
		// 进程已不在，缓存必然悬空，清理后重建
		b.invalidateNamespace()
	}

	// 2. 冷启动：获取 Application（GetActiveOutlook 在 COM 线程上内联执行）
	app, err := b.GetActiveOutlook()
	if err != nil {
		return nil, err
	}
	// 注意：app 的所有权转移给缓存，此处不再 defer SafeRelease(app)

	// 3. 获取 Namespace
	nsVar, err := comutil.SafeCallMethod(app, "GetNamespace", "MAPI")
	if err != nil {
		comutil.SafeRelease(app)
		return nil, err
	}
	ns := nsVar.ToIDispatch()
	ns.AddRef()   // +1：缓存自身持有的引用
	nsVar.Clear() // -1：释放 VARIANT 持有的引用（净效果：缓存独占 1 份）

	// 4. 写入缓存
	b.cachedApp = app
	b.cachedNS = ns

	// 5. 返回调用方一份独立引用
	b.cachedNS.AddRef()
	return b.cachedNS, nil
}

// EnsurePSTMounted 确保指定季度的 PST 文件已挂载
func (b *COMBridge) EnsurePSTMounted(quarter router.QuarterInfo, rootPath string) (*ole.IDispatch, error) {
	var rootFolder *ole.IDispatch
	err := b.Submit(func() error {
		expectedPath := quarter.PSTFilePath(rootPath)
		var err error
		rootFolder, err = b.EnsurePSTMountedByPath(expectedPath)
		if err != nil {
			return err
		}
		// 修改显示名称
		oleutil.PutProperty(rootFolder, "Name", quarter.DisplayName())
		return nil
	})
	return rootFolder, err
}

// EnsurePSTMountedByPath 确保指定路径的 PST 文件已挂载并返回其 RootFolder
func (b *COMBridge) EnsurePSTMountedByPath(expectedPath string) (*ole.IDispatch, error) {
	var rootFolderVarDispatch *ole.IDispatch
	err := b.Submit(func() error {
		ns, err := b.getNamespace()
		if err != nil {
			return err
		}
		defer comutil.SafeRelease(ns)

		storesVar, err := comutil.SafeGetProperty(ns, "Stores")
		if err != nil {
			return err
		}
		defer storesVar.Clear()
		stores := storesVar.ToIDispatch()
		defer comutil.SafeRelease(stores)

		countVar, err := comutil.SafeGetProperty(stores, "Count")
		if err != nil {
			return err
		}
		count := int(countVar.Val)

		var targetStore *ole.IDispatch

		for i := 1; i <= count; i++ {
			storeVar, err := comutil.SafeCallMethod(stores, "Item", i)
			if err != nil {
				continue
			}
			store := storeVar.ToIDispatch()
			pathVar, err := comutil.SafeGetProperty(store, "FilePath")
			if err == nil && pathVar.Value() != nil {
				path := pathVar.ToString()
				pathVar.Clear()
				if strings.EqualFold(path, expectedPath) {
					targetStore = store
					targetStore.AddRef()
					storeVar.Clear()
					break // Found it
				}
			}
			storeVar.Clear()
		}

		if targetStore == nil {
			// 未挂载，尝试 AddStoreEx (olStoreUnicode = 2)
			_, err := comutil.SafeCallMethod(ns, "AddStoreEx", expectedPath, 2)
			if err != nil {
				_, err = comutil.SafeCallMethod(ns, "AddStore", expectedPath)
				if err != nil {
					return fmt.Errorf("failed to mount PST %s: %w", expectedPath, err)
				}
			}

			// 重新查找刚挂载的 store
			countVar, _ = comutil.SafeGetProperty(stores, "Count")
			count = int(countVar.Val)
			for i := 1; i <= count; i++ {
				storeVar, err := comutil.SafeCallMethod(stores, "Item", i)
				if err != nil {
					continue
				}
				store := storeVar.ToIDispatch()
				pathVar, err := comutil.SafeGetProperty(store, "FilePath")
				if err == nil && pathVar.Value() != nil {
					isMatch := strings.EqualFold(pathVar.ToString(), expectedPath)
					pathVar.Clear()
					if isMatch {
						targetStore = store
						targetStore.AddRef()
						storeVar.Clear()
						break
					}
				}
				storeVar.Clear()
			}
		}

		if targetStore == nil {
			return fmt.Errorf("store mounted but not found in Stores collection: %s", expectedPath)
		}
		defer comutil.SafeRelease(targetStore)

		// 返回根文件夹
		rootFolderVar, err := comutil.SafeCallMethod(targetStore, "GetRootFolder")
		if err != nil {
			return fmt.Errorf("failed to get root folder: %w", err)
		}
		rootFolderVarDispatch = rootFolderVar.ToIDispatch()
		return nil
	})
	return rootFolderVarDispatch, err
}

// EnsurePSTFolder 确保 PST 内指定路径的文件夹存在
func (b *COMBridge) EnsurePSTFolder(pstRoot *ole.IDispatch, folderPath string) (*ole.IDispatch, error) {
	var resultFolder *ole.IDispatch
	err := b.Submit(func() error {
		if folderPath == "" {
			pstRoot.AddRef()
			resultFolder = pstRoot
			return nil
		}

		parts := strings.Split(strings.ReplaceAll(folderPath, "\\", "/"), "/")
		currentFolder := pstRoot
		currentFolder.AddRef()

		for _, part := range parts {
			if part == "" {
				continue
			}

			foldersVar, err := comutil.SafeGetProperty(currentFolder, "Folders")
			if err != nil {
				comutil.SafeRelease(currentFolder)
				return fmt.Errorf("failed to get Folders: %w", err)
			}
			folders := foldersVar.ToIDispatch()

			var nextFolder *ole.IDispatch
			// 尝试通过名称获取
			folderVar, err := comutil.SafeCallMethod(folders, "Item", part)
			if err == nil && folderVar.Value() != nil {
				nextFolder = folderVar.ToIDispatch()
				nextFolder.AddRef()
				folderVar.Clear()
			} else {
				// 不存在则创建
				newFolderVar, err := comutil.SafeCallMethod(folders, "Add", part)
				if err != nil {
					foldersVar.Clear()
					comutil.SafeRelease(currentFolder)
					return fmt.Errorf("failed to create folder %s: %w", part, err)
				}
				nextFolder = newFolderVar.ToIDispatch()
				nextFolder.AddRef()
				newFolderVar.Clear()
			}

			foldersVar.Clear()
			comutil.SafeRelease(currentFolder)
			currentFolder = nextFolder
		}

		resultFolder = currentFolder
		return nil
	})
	return resultFolder, err
}

// IsStoreMounted 通过物理路径判断 Store 是否已挂载
func (b *COMBridge) IsStoreMounted(filePath string) (bool, error) {
	psts, err := b.GetMountedPSTs()
	if err != nil {
		return false, err
	}
	for _, p := range psts {
		if strings.EqualFold(p, filePath) {
			return true, nil
		}
	}
	return false, nil
}

// GetMountedPSTs 返回当前已挂载的所有 PST 文件路径
func (b *COMBridge) GetMountedPSTs() ([]string, error) {
	var paths []string
	err := b.Submit(func() error {
		ns, err := b.getNamespace()
		if err != nil {
			return err
		}
		defer comutil.SafeRelease(ns)

		storesVar, err := comutil.SafeGetProperty(ns, "Stores")
		if err != nil {
			return err
		}
		defer storesVar.Clear()
		stores := storesVar.ToIDispatch()
		defer comutil.SafeRelease(stores)

		countVar, err := comutil.SafeGetProperty(stores, "Count")
		if err != nil {
			return err
		}
		count := int(countVar.Val)
		countVar.Clear()

		for i := 1; i <= count; i++ {
			storeVar, err := comutil.SafeCallMethod(stores, "Item", i)
			if err != nil {
				continue
			}
			store := storeVar.ToIDispatch()
			pathVar, err := comutil.SafeGetProperty(store, "FilePath")
			if err == nil && pathVar.Value() != nil {
				path := pathVar.ToString()
				if strings.HasSuffix(strings.ToLower(path), ".pst") {
					paths = append(paths, path)
				}
			}
			if pathVar != nil {
				pathVar.Clear()
			}
			storeVar.Clear()
		}
		return nil
	})
	return paths, err
}

// RemoveStore 卸载指定的 PST (通过其 RootFolder)
func (b *COMBridge) RemoveStore(rootFolder *ole.IDispatch) error {
	return b.Submit(func() error {
		ns, err := b.getNamespace()
		if err != nil {
			return err
		}
		defer comutil.SafeRelease(ns)

		_, err = comutil.SafeCallMethod(ns, "RemoveStore", rootFolder)
		return err
	})
}
