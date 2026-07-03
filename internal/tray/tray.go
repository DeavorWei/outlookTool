package tray

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"time"

	"github.com/getlantern/systray"

	"outlook-archiver/internal/config"
	"outlook-archiver/internal/logger"
	"outlook-archiver/internal/monitor"
	"go.uber.org/zap"
	"outlook-archiver/internal/registry"
	"outlook-archiver/internal/scheduler"
)

// normalIcon 使用简单的 1x1 透明 ICO 字节数组模拟，后续可替换为真实的托盘图标。
var normalIcon = []byte{
	0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00,
	0x18, 0x00, 0x30, 0x00, 0x00, 0x00, 0x16, 0x00, 0x00, 0x00, 0x28, 0x00,
	0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x01, 0x00,
	0x18, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
	0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
	0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
}

var (
	workingIcon = normalIcon // 在完整版本中可使用不同的“旋转图标”代替
	warningIcon = normalIcon
	errorIcon   = normalIcon
)

var (
	trayCtx    context.Context
	trayCancel context.CancelFunc
)

// InitTray 初始化并运行系统托盘
func InitTray(sched *scheduler.Scheduler, cfg *config.Config, zlog *zap.Logger) {
	trayCtx, trayCancel = context.WithCancel(context.Background())
	systray.Run(func() {
		onReady(sched, cfg, zlog)
	}, onExit)
}

func onReady(sched *scheduler.Scheduler, cfg *config.Config, zlog *zap.Logger) {
	systray.SetIcon(normalIcon)
	systray.SetTooltip("Outlook Auto-Archiver - 运行中")

	mArchiveOnce := systray.AddMenuItem("立即执行一次", "手动触发常规归档")
	mReorganize := systray.AddMenuItem("全量整理", "全量归档+PST纠偏")
	systray.AddSeparator()
	mOpenLog := systray.AddMenuItem("打开日志目录", "")
	mOpenConfig := systray.AddMenuItem("打开配置文件", "")
	mReloadConfig := systray.AddMenuItem("重新加载配置", "修改配置文件后点击生效")

	// 检查当前注册表状态
	autoStartEnabled := false
	// 由于模块可能还未完全实现，如果有panic可以通过 defer 捕获或者在此假设可用
	// 此处正常调用 registry 的方法
	autoStartEnabled = registry.IsAutoStartEnabled()
	mAutoStart := systray.AddMenuItemCheckbox("开机自启", "", autoStartEnabled)

	systray.AddSeparator()
	mQuit := systray.AddMenuItem("退出", "")

	// 启动 goroutine 定期监控 scheduler 状态以更新 UI (图标及菜单禁用状态)
	go func() {
		ticker := time.NewTicker(1 * time.Second)
		defer ticker.Stop()
		for {
			select {
			case <-trayCtx.Done():
				return
			case <-ticker.C:
				if sched == nil {
					continue
				}
				state := sched.GetState()
				safeCfg := sched.GetConfigCopy()
				diskStatus, _ := monitor.CheckDiskSpace(safeCfg.PSTRootPath)
				switch state {
				case scheduler.StateIdle, scheduler.StatePaused:
					if diskStatus == monitor.DiskCritical {
						systray.SetIcon(errorIcon)
						systray.SetTooltip("Outlook Auto-Archiver - 磁盘空间不足")
					} else if diskStatus == monitor.DiskWarning {
						systray.SetIcon(warningIcon)
						systray.SetTooltip("Outlook Auto-Archiver - 磁盘空间警告")
					} else {
						systray.SetIcon(normalIcon)
						systray.SetTooltip("Outlook Auto-Archiver - 运行中")
					}
					mArchiveOnce.Enable()
					mReorganize.Enable()
					mReloadConfig.Enable()
				case scheduler.StateArchiving:
					if diskStatus == monitor.DiskCritical {
						systray.SetIcon(errorIcon)
					} else if diskStatus == monitor.DiskWarning {
						systray.SetIcon(warningIcon)
					} else {
						systray.SetIcon(normalIcon)
					}
					systray.SetTooltip("归档中...")
					mArchiveOnce.Disable()
					mReorganize.Enable()
					mReloadConfig.Disable()
				case scheduler.StateReorganizing:
					systray.SetIcon(workingIcon)
					// Tooltip 会在 TriggerReorganize 的 progressCb 中实时更新，这里不覆盖。
					mArchiveOnce.Disable()
					mReorganize.Disable()
					mReloadConfig.Disable()
				}
			}
		}
	}()

	// 响应用户的点击事件
	go func() {
		for {
			select {
			case <-trayCtx.Done():
				return
			case <-mArchiveOnce.ClickedCh:
				if sched != nil {
					_ = sched.TriggerOnce(trayCtx)
				}
			case <-mReorganize.ClickedCh:
				if sched != nil {
					// 触发全量整理，并传入进度更新回调
					_ = sched.TriggerReorganize(trayCtx, func(info scheduler.ProgressInfo) {
						systray.SetTooltip(fmt.Sprintf("全量整理中... 阶段%d - 已纠偏 %d 封", info.Phase, info.Processed))
					})
				}
			case <-mOpenLog.ClickedCh:
				if logger.CurrentLogDir != "" {
					exec.Command("explorer", logger.CurrentLogDir).Start()
				}
			case <-mOpenConfig.ClickedCh:
				exeDir, err := os.Executable()
				if err == nil {
					configPath := filepath.Join(filepath.Dir(exeDir), "config.yaml")
					exec.Command("cmd", "/c", "start", "", configPath).Start()
				}
			case <-mReloadConfig.ClickedCh:
				if sched != nil {
					exeDir, err := os.Executable()
					if err == nil {
						configPath := filepath.Join(filepath.Dir(exeDir), "config.yaml")
						err = sched.ReloadConfig(configPath)
						if err != nil {
							zlog.Error("重新加载配置失败", zap.Error(err))
							systray.SetTooltip("Outlook Auto-Archiver - 加载配置失败")
						} else {
							zlog.Info("重新加载配置成功")
						}
					}
				}
			case <-mAutoStart.ClickedCh:
				if mAutoStart.Checked() {
					err := registry.SetAutoStart(false)
					if err == nil {
						mAutoStart.Uncheck()
						zlog.Info("已取消开机自启")
					} else {
						zlog.Error("取消开机自启失败", zap.Error(err))
					}
				} else {
					err := registry.SetAutoStart(true)
					if err == nil {
						mAutoStart.Check()
						zlog.Info("已开启开机自启")
					} else {
						zlog.Error("开启开机自启失败", zap.Error(err))
					}
				}
			case <-mQuit.ClickedCh:
				if sched != nil && sched.GetState() == scheduler.StateReorganizing {
					zlog.Warn("全量整理中，禁止退出")
					continue
				}
				systray.Quit()
			}
		}
	}()
}

func onExit() {
	if trayCancel != nil {
		trayCancel()
	}
	// 系统托盘退出时的清理逻辑（如停止调度器等，可根据实际应用周期进行）
}
