package registry

import (
	"os"
	"path/filepath"
	"strings"

	"golang.org/x/sys/windows/registry"
)

const (
	runKeyPath = `Software\Microsoft\Windows\CurrentVersion\Run`
	appName    = "OutlookAutoArchiver"
)

// SetAutoStart 写入/删除开机自启注册表项
// HKCU\Software\Microsoft\Windows\CurrentVersion\Run
func SetAutoStart(enabled bool) error {
	k, err := registry.OpenKey(registry.CURRENT_USER, runKeyPath, registry.SET_VALUE)
	if err != nil {
		return err
	}
	defer k.Close()

	if enabled {
		exePath, err := os.Executable()
		if err != nil {
			return err
		}
		// 使用双引号包裹路径，防止路径中有空格
		exePath = `"` + filepath.Clean(exePath) + `"`
		return k.SetStringValue(appName, exePath)
	} else {
		err := k.DeleteValue(appName)
		if err != nil && err != registry.ErrNotExist {
			return err
		}
		return nil
	}
}

// IsAutoStartEnabled 检查当前是否已启用自启
func IsAutoStartEnabled() bool {
	k, err := registry.OpenKey(registry.CURRENT_USER, runKeyPath, registry.QUERY_VALUE)
	if err != nil {
		return false
	}
	defer k.Close()

	val, _, err := k.GetStringValue(appName)
	if err != nil {
		return false
	}

	exePath, err := os.Executable()
	if err != nil {
		return val != ""
	}

	cleanPath := `"` + filepath.Clean(exePath) + `"`
	// 忽略大小写比较路径
	return strings.EqualFold(val, cleanPath)
}
