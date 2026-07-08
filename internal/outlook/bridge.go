package outlook

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"runtime"
	"strings"
	"sync/atomic"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"go.uber.org/zap"
	"golang.org/x/sys/windows"

	"outlook-archiver/pkg/comutil"
)

type comTask struct {
	fn    func() error
	errCh chan error
}

// COMBridge 封装 COM 操作的桥接层
type COMBridge struct {
	taskCh   chan comTask // 所有 COM 操作通过此 channel 提交
	threadID uint32
	logger   *zap.Logger

	// MAPI Namespace 单例缓存。
	// 约束：仅在 COM 线程访问（所有读写均位于 b.Submit 闭包内），无需加锁。
	// 缓存自身持有一份引用，调用方通过 AddRef 获得独立引用。
	cachedApp *ole.IDispatch // Outlook.Application 的 IDispatch
	cachedNS  *ole.IDispatch // Namespace("MAPI") 的 IDispatch
}

// NewCOMBridge 创建一个新的 COMBridge
func NewCOMBridge(logger *zap.Logger) *COMBridge {
	return &COMBridge{
		taskCh: make(chan comTask, 64),
		logger: logger,
	}
}

// Run 启动 COM 工作线程（必须在独立 goroutine 中调用）
func (b *COMBridge) Run(ctx context.Context, ready chan<- struct{}) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	atomic.StoreUint32(&b.threadID, windows.GetCurrentThreadId())

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	defer ole.CoUninitialize()

	if b.logger != nil {
		b.logger.Info("COM 桥接工作线程已就绪，等待执行底层任务...")
	}
	close(ready)

	defer func() {
		if b.logger != nil {
			b.logger.Info("COM 桥接工作线程正在安全退出...")
		}
		// 释放缓存的 COM 对象（必须在 CoUninitialize 之前，此时 COM 仍可用）
		b.invalidateNamespace()

		// 排空 taskCh，让等待的 Submit 收到错误
		for {
			select {
			case task := <-b.taskCh:
				if task.errCh != nil {
					task.errCh <- fmt.Errorf("COM bridge is shutting down")
				}
			default:
				return
			}
		}
	}()

	for {
		select {
		case task := <-b.taskCh:
			func() {
				defer func() {
					if r := recover(); r != nil {
						err := fmt.Errorf("COM 任务 panic: %v", r)
						// 同时写入应用日志便于排查（stderr 保留为兜底）
						if b.logger != nil {
							b.logger.Error("COM 任务出现 panic", zap.Any("panic", r), zap.Stack("stack"))
						}
						fmt.Fprintf(os.Stderr, "%s\n", err)
						if task.errCh != nil {
							task.errCh <- err
						}
					}
				}()
				err := task.fn() // 在锁定的 OS 线程上执行
				if task.errCh != nil {
					task.errCh <- err
				}
			}()
		case <-ctx.Done():
			return
		}
	}
}

// Submit 向 COM 线程提交操作并等待结果
func (b *COMBridge) Submit(fn func() error) error {
	return b.SubmitWithContext(context.Background(), fn)
}

// SubmitWithContext 向 COM 线程提交操作并支持上下文取消
func (b *COMBridge) SubmitWithContext(ctx context.Context, fn func() error) error {
	if tid := atomic.LoadUint32(&b.threadID); tid != 0 && windows.GetCurrentThreadId() == tid {
		return fn()
	}

	errCh := make(chan error, 1)
	task := comTask{
		fn:    fn,
		errCh: errCh,
	}

	select {
	case b.taskCh <- task:
	case <-ctx.Done():
		return ctx.Err()
	}

	select {
	case err := <-errCh:
		return err
	case <-ctx.Done():
		return ctx.Err()
	}
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
		// 即将关闭 Outlook，预先令 Namespace 缓存失效，避免悬空代理
		b.invalidateNamespace()

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

// invalidateNamespace 释放并清空缓存的 Application 与 Namespace。
//
// 必须在 COM 线程调用。调用后下一次 getNamespace 会重新走冷启动路径。
// 幂等：对 nil 字段调用 SafeRelease 是安全的。
func (b *COMBridge) invalidateNamespace() {
	if b.cachedNS != nil {
		comutil.SafeRelease(b.cachedNS)
		b.cachedNS = nil
	}
	if b.cachedApp != nil {
		comutil.SafeRelease(b.cachedApp)
		b.cachedApp = nil
	}
}

// InvalidateNamespaceCache 清空 MAPI Namespace 缓存（外部调用入口）。
//
// 供非 COM 线程的调用方在 Outlook 进程即将被重启时调用。
// 通过 Submit 派发到 COM 线程执行，保证缓存字段的读写线程安全。
func (b *COMBridge) InvalidateNamespaceCache() {
	_ = b.Submit(func() error {
		b.invalidateNamespace()
		return nil
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
