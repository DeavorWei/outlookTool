package outlook

import (
	"fmt"
	"time"

	"github.com/go-ole/go-ole"
	"outlook-archiver/pkg/comutil"
)

// MoveItem 移动邮件到目标文件夹
func (b *COMBridge) MoveItem(mailItem *ole.IDispatch, targetFolder *ole.IDispatch) error {
	return b.Submit(func() error {
		_, err := comutil.SafeCallMethod(mailItem, "Move", targetFolder)
		return err
	})
}

// GetSubject 获取邮件主题
func (b *COMBridge) GetSubject(mailItem *ole.IDispatch) string {
	var subject string
	_ = b.Submit(func() error {
		subjectVar, err := comutil.SafeGetProperty(mailItem, "Subject")
		if err != nil || subjectVar.Value() == nil {
			return err
		}
		defer subjectVar.Clear()
		subject = subjectVar.ToString()
		return nil
	})
	return subject
}

// GetMailTime 获取邮件的时间属性 (如 ReceivedTime 或 SentOn)
func (b *COMBridge) GetMailTime(mailItem *ole.IDispatch, timeField string) (time.Time, error) {
	var t time.Time
	err := b.Submit(func() error {
		timeVar, err := comutil.SafeGetProperty(mailItem, timeField)
		if err != nil || timeVar.Value() == nil {
			// 回退到 CreationTime
			timeVar, err = comutil.SafeGetProperty(mailItem, "CreationTime")
			if err != nil || timeVar.Value() == nil {
				return fmt.Errorf("failed to get time field %s and CreationTime", timeField)
			}
		}
		defer timeVar.Clear()
		t, err = parseTime(timeVar.Value())
		return err
	})
	return t, err
}

// parseTime 尝试将 COM 返回的类型转换为 time.Time
func parseTime(val interface{}) (time.Time, error) {
	switch v := val.(type) {
	case time.Time:
		return v, nil
	case float64: // OLE Automation Date
		days := int(v)
		fraction := v - float64(days)
		t := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).AddDate(0, 0, days)
		t = t.Add(time.Duration(fraction * 24 * float64(time.Hour)))
		return t, nil
	default:
		return time.Time{}, fmt.Errorf("unsupported time format: %T", val)
	}
}

// GetItemClass 获取邮件的 MessageClass
func (b *COMBridge) GetItemClass(mailItem *ole.IDispatch) string {
	var class string
	_ = b.Submit(func() error {
		classVar, err := comutil.SafeGetProperty(mailItem, "MessageClass")
		if err != nil || classVar.Value() == nil {
			return err
		}
		defer classVar.Clear()
		class = classVar.ToString()
		return nil
	})
	return class
}
