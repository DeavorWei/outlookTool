package comutil

import (
	"fmt"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// COMRef 包装 COM 对象引用，确保释放
type COMRef struct {
	disp *ole.IDispatch
	name string // 用于调试日志
}

// NewCOMRef 创建一个新的 COM 引用对象
func NewCOMRef(disp *ole.IDispatch, name string) *COMRef {
	return &COMRef{disp: disp, name: name}
}

// Release 安全释放包装的 COM 对象
func (r *COMRef) Release() {
	if r != nil && r.disp != nil {
		r.disp.Release()
		r.disp = nil
	}
}

// Dispatch 返回内部的 IDispatch 指针
func (r *COMRef) Dispatch() *ole.IDispatch {
	if r == nil {
		return nil
	}
	return r.disp
}

// SafeRelease 安全释放 COM 对象，带 nil 检查和 panic recovery
func SafeRelease(disp *ole.IDispatch) {
	defer func() {
		if r := recover(); r != nil {
			// recover from panic during release
			fmt.Printf("panic in SafeRelease: %v\n", r)
		}
	}()
	if disp != nil {
		disp.Release()
	}
}

// SafeCallMethod 调用 COM 方法，自动处理错误码分类
func SafeCallMethod(disp *ole.IDispatch, method string, params ...interface{}) (*ole.VARIANT, error) {
	if disp == nil {
		return nil, fmt.Errorf("nil IDispatch for method %s", method)
	}
	v, err := oleutil.CallMethod(disp, method, params...)
	if err != nil {
		return nil, fmt.Errorf("CallMethod %s failed: %w", method, err)
	}
	return v, nil
}

// SafeGetProperty 获取 COM 属性值
func SafeGetProperty(disp *ole.IDispatch, prop string) (*ole.VARIANT, error) {
	if disp == nil {
		return nil, fmt.Errorf("nil IDispatch for property %s", prop)
	}
	v, err := oleutil.GetProperty(disp, prop)
	if err != nil {
		return nil, fmt.Errorf("GetProperty %s failed: %w", prop, err)
	}
	return v, nil
}
