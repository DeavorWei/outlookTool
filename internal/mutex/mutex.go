package mutex

import (
	"fmt"
	"os/user"

	"golang.org/x/sys/windows"
)

// TryLock 尝试获取系统级互斥锁
// 命名格式: Global\OutlookAutoArchiver_{UserSID}
func TryLock() (release func(), err error) {
	u, err := user.Current()
	if err != nil {
		return nil, fmt.Errorf("failed to get current user: %w", err)
	}

	name := `Global\OutlookAutoArchiver_` + u.Uid
	namePtr, err := windows.UTF16PtrFromString(name)
	if err != nil {
		return nil, err
	}

	handle, err := windows.CreateMutex(nil, false, namePtr)
	if err != nil && err != windows.ERROR_ALREADY_EXISTS {
		return nil, fmt.Errorf("failed to create mutex: %w", err)
	}
	if err == windows.ERROR_ALREADY_EXISTS || windows.GetLastError() == windows.ERROR_ALREADY_EXISTS {
		if handle != 0 {
			windows.CloseHandle(handle)
		}
		return nil, fmt.Errorf("程序已在运行中 (already running)")
	}

	release = func() {
		if handle != 0 {
			windows.ReleaseMutex(handle)
			windows.CloseHandle(handle)
		}
	}
	return release, nil
}
