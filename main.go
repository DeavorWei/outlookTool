package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"outlook-archiver/internal/archiver"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/logger"
	"outlook-archiver/internal/mutex"
	"outlook-archiver/internal/outlook"
	"outlook-archiver/internal/scheduler"
	"outlook-archiver/internal/tray"
)

func main() {
	// 1. 获取系统级单例锁
	release, err := mutex.TryLock()
	if err != nil {
		fmt.Printf("程序已在运行或获取单例锁失败: %v\n", err)
		os.Exit(1)
	}
	defer release()

	// 2. 加载配置
	exePath, err := os.Executable()
	if err != nil {
		log.Fatalf("无法获取程序执行路径: %v", err)
	}
	configPath := filepath.Join(filepath.Dir(exePath), "config.yaml")
	cfg, isFirstRun, err := config.LoadConfig(configPath)
	if err != nil {
		log.Fatalf("加载配置失败: %v", err)
	}

	// 3. 初始化日志
	zlog, err := logger.InitLogger(cfg)
	if err != nil {
		log.Fatalf("初始化日志失败: %v", err)
	}
	defer zlog.Sync()
	zlog.Info("Outlook Auto-Archiver 启动中...")

	// 4. 初始化 COM 桥接层
	comBridge := outlook.NewCOMBridge(zlog)
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()

	// 启动 COM 线程并等待就绪
	ready := make(chan struct{})
	go comBridge.Run(ctx, ready)
	<-ready

	// 5. 初始化核心引擎
	// 配置提供方：实时返回调度器当前生效配置（支持热重载）。
	// 通过闭包引用 sched，需在调度器构造后再赋值（见下方），归档/整理逻辑即可拿到最新配置。
	var sched *scheduler.Scheduler
	cfgProvider := func() config.Config {
		if sched == nil {
			return config.Config{}
		}
		return sched.GetConfigCopy()
	}
	arc := archiver.NewArchiver(cfgProvider, comBridge, zlog)
	reorg := archiver.NewReorganizer(cfgProvider, comBridge, arc, zlog)

	// 6. 初始化调度器
	sched = scheduler.NewScheduler(cfg, arc, reorg, zlog)
	sched.Start(ctx)
	defer sched.Stop()

	// 7. 初始化并运行系统托盘 (会阻塞主线程)
	tray.InitTray(sched, cfg, zlog, isFirstRun)

	zlog.Info("正在停止后台任务...")
	cancel()     // 取消 context，通知各组件停止
	sched.Stop() // 显式停止调度器

	// 等待正在进行的任务结束 (最多等待 5 秒)
	waitCh := make(chan struct{})
	go func() {
		for sched.GetState() != scheduler.StateIdle {
			time.Sleep(100 * time.Millisecond)
		}
		close(waitCh)
	}()

	select {
	case <-waitCh:
		zlog.Info("所有后台任务已优雅完成")
	case <-time.After(15 * time.Second):
		zlog.Warn("等待后台任务超时，强制退出")
	}

	zlog.Info("Outlook Auto-Archiver 已退出")
}
