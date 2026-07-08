package monitor

import (
	"golang.org/x/sys/windows"
	"os"
)

type DiskStatus int

const (
	DiskOK       DiskStatus = iota
	DiskWarning             // < 1GB
	DiskCritical            // < 500MB
)

type PSTSizeStatus int

const (
	PSTSizeOK       PSTSizeStatus = iota
	PSTSizeWarning                // > 15GB
	PSTSizeCritical               // > 20GB
)

// CheckDiskSpace 检查目标目录可用空间
func CheckDiskSpace(path string) (DiskStatus, error) {
	var freeBytesAvailable, totalNumberOfBytes, totalNumberOfFreeBytes uint64
	pathPtr, err := windows.UTF16PtrFromString(path)
	if err != nil {
		return DiskCritical, err
	}
	err = windows.GetDiskFreeSpaceEx(pathPtr, &freeBytesAvailable, &totalNumberOfBytes, &totalNumberOfFreeBytes)
	if err != nil {
		return DiskCritical, err
	}

	const GB = 1024 * 1024 * 1024
	const MB = 1024 * 1024
	if freeBytesAvailable < 500*MB {
		return DiskCritical, nil
	}
	if freeBytesAvailable < 1*GB {
		return DiskWarning, nil
	}
	return DiskOK, nil
}

// CheckPSTSize 检查单个 PST 文件大小
func CheckPSTSize(pstPath string) (PSTSizeStatus, error) {
	info, err := os.Stat(pstPath)
	if err != nil {
		return PSTSizeCritical, err
	}
	size := info.Size()

	const GB = 1024 * 1024 * 1024
	if size > 20*GB {
		return PSTSizeCritical, nil
	}
	if size > 15*GB {
		return PSTSizeWarning, nil
	}
	return PSTSizeOK, nil
}
