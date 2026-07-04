package outlook

import (
	"context"
	"fmt"
	"os/exec"
	"runtime"
	"strings"
	"sync/atomic"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/sys/windows"
)

// COMBridge 封装 COM 操作的桥接层
type COMBridge struct {
	taskCh   chan func() // 所有 COM 操作通过此 channel 提交
	resultCh chan error
	threadID uint32
}

// NewCOMBridge 创建一个新的 COMBridge
func NewCOMBridge() *COMBridge {
	return &COMBridge{
		taskCh:   make(chan func(), 64),
		resultCh: make(chan error),
	}
}

// Run 启动 COM 工作线程（必须在独立 goroutine 中调用）
func (b *COMBridge) Run(ctx context.Context) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	
	atomic.StoreUint32(&b.threadID, windows.GetCurrentThreadId())

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	defer ole.CoUninitialize()

	for {
		select {
		case task := <-b.taskCh:
			task() // 在锁定的 OS 线程上执行
		case <-ctx.Done():
			return
		}
	}
}

// Submit 向 COM 线程提交操作并等待结果
func (b *COMBridge) Submit(fn func() error) error {
	if tid := atomic.LoadUint32(&b.threadID); tid != 0 && windows.GetCurrentThreadId() == tid {
		return fn()
	}

	// 使用局部 channel 接收结果，保证并发安全
	errCh := make(chan error, 1)
	b.taskCh <- func() {
		errCh <- fn()
	}
	return <-errCh
}

// IsOutlookRunning 通过 Windows API 扫描进程列表
func IsOutlookRunning() bool {
	snapshot, err := windows.CreateToolhelp32Snapshot(windows.TH32CS_SNAPPROCESS, 0)
	if err != nil {
		return false
	}
	defer windows.CloseHandle(snapshot)

	var pe32 windows.ProcessEntry32
	pe32.Size = uint32(unsafe.Sizeof(pe32))

	err = windows.Process32First(snapshot, &pe32)
	for err == nil {
		exeName := windows.UTF16ToString(pe32.ExeFile[:])
		if strings.EqualFold(exeName, "OUTLOOK.EXE") {
			return true
		}
		err = windows.Process32Next(snapshot, &pe32)
	}
	return false
}

// GetActiveOutlook 获取当前运行的 Outlook 实例
// 返回 Application 的 IDispatch 指针
func (b *COMBridge) GetActiveOutlook() (*ole.IDispatch, error) {
	var disp *ole.IDispatch
	var err error

	err = b.Submit(func() error {
		unknown, errGet := oleutil.GetActiveObject("Outlook.Application")
		if errGet != nil {
			return fmt.Errorf("请确认Outlook已打开且未以管理员身份运行 (错误原因为: %w)", errGet)
		}

		disp, errGet = unknown.QueryInterface(ole.IID_IDispatch)
		unknown.Release() // 释放 IUnknown
		if errGet != nil {
			return fmt.Errorf("failed to query IDispatch: %w", errGet)
		}

		return nil
	})

	if err != nil {
		return nil, err
	}
	return disp, nil
}

// QuitOutlook 尝试通过 COM 接口优雅关闭 Outlook
func (b *COMBridge) QuitOutlook() error {
	return b.Submit(func() error {
		unknown, errGet := oleutil.GetActiveObject("Outlook.Application")
		if errGet != nil {
			return nil // 已经没有运行，或者无法获取，则视作已经关闭
		}
		disp, errQuery := unknown.QueryInterface(ole.IID_IDispatch)
		unknown.Release()
		if errQuery != nil {
			return errQuery
		}
		defer disp.Release()

		_, errCall := oleutil.CallMethod(disp, "Quit")
		return errCall
	})
}

// ForceKillOutlook 强制结束 Outlook 进程
func ForceKillOutlook() error {
	cmd := exec.Command("taskkill", "/F", "/IM", "OUTLOOK.EXE")
	return cmd.Run()
}

// StartOutlook 通过命令行启动 Outlook
func StartOutlook() error {
	cmd := exec.Command("cmd", "/c", "start", "outlook")
	return cmd.Start()
}

