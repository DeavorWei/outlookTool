package scheduler

import (
	"context"
	"fmt"
	"sync"
	"sync/atomic"
	"time"

	"go.uber.org/zap"

	"outlook-archiver/internal/archiver"
	"outlook-archiver/internal/config"
	"outlook-archiver/internal/monitor"
)

type SchedulerState int

const (
	StateIdle SchedulerState = iota
	StateArchiving
	StateReorganizing
	StateRestoring
)

// ProgressInfo 进度信息（供托盘 Tooltip 显示）
type ProgressInfo struct {
	Phase      int // 1 或 2
	Processed  int
	Rectified  int
	CurrentPST string
}

// ConfigProvider 允许安全获取当前生效的配置副本
func (s *Scheduler) GetConfigCopy() config.Config {
	val := s.cfgVal.Load()
	if cfg, ok := val.(*config.Config); ok && cfg != nil {
		return *cfg
	}
	return config.Config{}
}

// GetMountedPSTs delegates to Archiver to fetch mounted PSTs
func (s *Scheduler) GetMountedPSTs() ([]string, error) {
	return s.archiver.GetMountedPSTs()
}

// ReloadConfig 从文件重新加载配置
func (s *Scheduler) ReloadConfig(path string) error {
	s.mu.Lock()
	if s.state != StateIdle {
		s.mu.Unlock()
		return fmt.Errorf("当前状态不可重新加载配置，请在空闲时重试")
	}
	s.mu.Unlock() // 检查状态后释放，避免 IO 阻塞

	newCfg, _, err := config.LoadConfig(path)
	if err != nil {
		return err
	}

	// 原子替换
	s.cfgVal.Store(newCfg)

	s.mu.Lock()
	defer s.mu.Unlock()

	// 更新日志级别等如果需要的话（暂时略过复杂的 logger 热更）

	// 重置定时器间隔
	if s.ticker != nil {
		interval := time.Duration(newCfg.PollIntervalMin) * time.Minute
		if interval <= 0 {
			interval = 60 * time.Minute
		}
		s.ticker.Reset(interval)
	}

	s.logger.Info("配置文件重新加载成功")
	return nil
}

type Scheduler struct {
	cfgVal   atomic.Value
	archiver *archiver.Archiver
	reorg    *archiver.Reorganizer
	logger   *zap.Logger
	ticker   *time.Ticker
	state    SchedulerState
	mu       sync.Mutex
	stopCh   chan struct{}
}

func NewScheduler(cfg *config.Config, arc *archiver.Archiver, reorg *archiver.Reorganizer, logger *zap.Logger) *Scheduler {
	s := &Scheduler{
		archiver: arc,
		reorg:    reorg,
		logger:   logger,
		state:    StateIdle,
		stopCh:   make(chan struct{}),
	}
	s.cfgVal.Store(cfg)
	return s
}

// Start 启动定时调度
func (s *Scheduler) Start(ctx context.Context) {
	s.mu.Lock()
	if s.ticker != nil {
		s.mu.Unlock()
		return // already started
	}
	cfg := s.GetConfigCopy()
	interval := time.Duration(cfg.PollIntervalMin) * time.Minute
	if interval <= 0 {
		interval = 60 * time.Minute // default fallback
	}
	s.ticker = time.NewTicker(interval)
	ticker := s.ticker
	s.stopCh = make(chan struct{})
	s.mu.Unlock()

	s.logger.Info("调度器已启动", zap.Duration("interval", interval))

	go func() {
		for {
			select {
			case <-ctx.Done():
				s.Stop()
				return
			case <-s.stopCh:
				return
			case <-ticker.C:
				_ = s.TriggerOnce(ctx)
			}
		}
	}()
}

// Stop 停止调度器
func (s *Scheduler) Stop() {
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.ticker != nil {
		s.ticker.Stop()
		s.ticker = nil
		close(s.stopCh)
		s.logger.Info("调度器已停止")
	}
}

// TriggerOnce 手动触发一次常规归档
func (s *Scheduler) TriggerOnce(ctx context.Context) error {
	s.mu.Lock()
	if s.state != StateIdle {
		s.mu.Unlock()
		return fmt.Errorf("当前状态不可触发归档: %v", s.state)
	}
	s.state = StateArchiving
	s.mu.Unlock()

	defer func() {
		s.mu.Lock()
		s.state = StateIdle
		s.mu.Unlock()
	}()

	s.logger.Info("开始常规归档任务")

	cfg := s.GetConfigCopy()
	// 检查磁盘空间
	status, err := monitor.CheckDiskSpace(cfg.PSTRootPath)
	if err != nil {
		s.logger.Error("检查磁盘空间失败", zap.Error(err))
		return err
	}
	if status == monitor.DiskCritical {
		err := fmt.Errorf("磁盘空间极度不足")
		s.logger.Error(err.Error())
		return err
	} else if status == monitor.DiskWarning {
		s.logger.Warn("磁盘空间不足 1GB")
	}

	opts := archiver.ArchiveOptions{
		MaxBatchSize: cfg.MaxBatchSize,
		SafeDelay:    time.Duration(cfg.SafeDelayMin) * time.Minute,
		RetainDays:   cfg.RetainDays,
		MoveInterval: time.Duration(cfg.MoveIntervalMs) * time.Millisecond,
		DryRun:       cfg.DryRun,
		CopyOnly:     cfg.CopyOnly,
	}

	res, err := s.archiver.Archive(ctx, opts)
	if err != nil {
		s.logger.Error("归档失败", zap.Error(err))
		return err
	}

	s.logger.Info("常规归档任务完成", zap.Any("result", res))

	return nil
}

// TriggerReorganize 触发全量整理（暂停定时器 → 执行 → 恢复）
func (s *Scheduler) TriggerReorganize(ctx context.Context, progressCb func(ProgressInfo)) error {
	s.mu.Lock()
	if s.state != StateIdle {
		s.mu.Unlock()
		return fmt.Errorf("当前状态不可触发全量整理: %v", s.state)
	}
	s.state = StateReorganizing

	// 暂停定时器
	if s.ticker != nil {
		s.ticker.Stop()
	}
	s.mu.Unlock()

	defer func() {
		s.mu.Lock()
		s.state = StateIdle
		if s.ticker != nil {
			// 恢复定时器
			cfg := s.GetConfigCopy()
			interval := time.Duration(cfg.PollIntervalMin) * time.Minute
			if interval <= 0 {
				interval = 60 * time.Minute
			}
			s.ticker.Reset(interval)
		}
		s.mu.Unlock()
	}()

	s.logger.Info("开始全量整理任务")

	cfg := s.GetConfigCopy()
	// 检查磁盘空间
	status, err := monitor.CheckDiskSpace(cfg.PSTRootPath)
	if err != nil {
		s.logger.Error("检查磁盘空间失败", zap.Error(err))
		return err
	}
	if status == monitor.DiskCritical {
		err := fmt.Errorf("磁盘空间极度不足")
		s.logger.Error(err.Error())
		return err
	} else if status == monitor.DiskWarning {
		s.logger.Warn("磁盘空间不足 1GB")
	}

	res, err := s.reorg.Reorganize(ctx, func(phase, processed, rectified int, currentPST string) {
		if progressCb != nil {
			progressCb(ProgressInfo{
				Phase:      phase,
				Processed:  processed,
				Rectified:  rectified,
				CurrentPST: currentPST,
			})
		}
	})
	if err != nil {
		s.logger.Error("全量整理失败", zap.Error(err))
		return err
	}

	s.logger.Info("全量整理任务完成", zap.Any("result", res))
	return nil
}

type RestoreRequest struct {
	DeleteEmptyPST   bool `json:"delete_empty_pst"`
	DeleteDuplicates bool `json:"delete_duplicates"`
}

// TriggerRestore 触发还原任务
func (s *Scheduler) TriggerRestore(ctx context.Context, req RestoreRequest) error {
	s.mu.Lock()
	if s.state != StateIdle {
		s.mu.Unlock()
		return fmt.Errorf("当前状态不可触发还原: %v", s.state)
	}
	s.state = StateRestoring

	// 暂停定时器
	if s.ticker != nil {
		s.ticker.Stop()
	}
	s.mu.Unlock()

	defer func() {
		s.mu.Lock()
		s.state = StateIdle
		if s.ticker != nil {
			cfg := s.GetConfigCopy()
			interval := time.Duration(cfg.PollIntervalMin) * time.Minute
			if interval <= 0 {
				interval = 60 * time.Minute
			}
			s.ticker.Reset(interval)
		}
		s.mu.Unlock()
	}()

	s.logger.Info("开始还原任务", zap.Bool("deleteEmpty", req.DeleteEmptyPST), zap.Bool("deleteDup", req.DeleteDuplicates))

	res, err := s.archiver.Restore(ctx, req.DeleteEmptyPST, req.DeleteDuplicates)
	if err != nil {
		s.logger.Error("还原任务失败", zap.Error(err))
		return err
	}

	s.logger.Info("还原任务完成", zap.Any("result", res))

	// 在数据还原任务结束后，立即执行独立的空文件清理流程
	s.archiver.ExecuteIndependentDeletion()

	return nil
}

// GetState 获取当前状态（供托盘 UI 查询）
func (s *Scheduler) GetState() SchedulerState {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.state
}
