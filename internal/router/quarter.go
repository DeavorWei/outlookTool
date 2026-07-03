package router

import (
	"fmt"
	"path/filepath"
	"regexp"
	"time"
)

// QuarterInfo 季度信息
type QuarterInfo struct {
	Year    int // e.g. 2024
	Quarter int // 1-4
}

// CalcQuarter 根据邮件时间计算所属季度
// 严禁传入 time.Now()，只接受邮件属性时间
func CalcQuarter(mailTime time.Time) QuarterInfo {
	return QuarterInfo{
		Year:    mailTime.Year(),
		Quarter: int(mailTime.Month()-1)/3 + 1,
	}
}

// PSTFileName 生成 PST 文件名：2024_Q2.pst
func (q QuarterInfo) PSTFileName() string {
	return fmt.Sprintf("%d_Q%d.pst", q.Year, q.Quarter)
}

// PSTFilePath 生成 PST 完整路径
func (q QuarterInfo) PSTFilePath(rootPath string) string {
	return filepath.Join(rootPath, q.PSTFileName())
}

// DisplayName 生成 Outlook 侧边栏显示名：2024_Q2
func (q QuarterInfo) DisplayName() string {
	return fmt.Sprintf("%d_Q%d", q.Year, q.Quarter)
}

var pstNameRegex = regexp.MustCompile(`(?i)^\d{4}_Q[1-4]\.pst$`)

// IsOurPSTName 判断文件名是否符合本工具命名规范
// 用于全量整理时区分本工具 PST 和第三方 PST
func IsOurPSTName(filename string) bool {
	return pstNameRegex.MatchString(filename)
}
