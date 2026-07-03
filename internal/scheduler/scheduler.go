package scheduler

import (
	"context"
	"fmt"
	"sync"
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
	StatePaused
)

// ProgressInfo 进度信息（供托盘 Tooltip 显示）
type ProgressInfo struct {
	Phase      int // 1 或 2
	Processed  int
	Rectified  int
	CurrentPST string
}

type Scheduler struct {
	cfg      *config.Config
	archiver *archiver.Archiver
	reorg    *archiver.Reorganizer
	logger   *zap.Logger
	ticker   *time.Ticker
	state    SchedulerState
	mu       sync.Mutex
	stopCh   chan struct{}
}

func NewScheduler(cfg *config.Config, arc *archiver.Archiver, reorg *archiver.Reorganizer, logger *zap.Logger) *Scheduler {
	return &Scheduler{
		cfg:      cfg,
		archiver: arc,
		reorg:    reorg,
		logger:   logger,
		state:    StateIdle,
		stopCh:   make(chan struct{}),
	}
}

// Start 启动定时调度
func (s *Scheduler) Start(ctx context.Context) {
	s.mu.Lock()
	if s.ticker != nil {
		s.mu.Unlock()
		return // already started
	}
	interval := time.Duration(s.cfg.PollIntervalMin) * time.Minute
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

	// 检查磁盘空间
	status, err := monitor.CheckDiskSpace(s.cfg.PSTRootPath)
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
		MaxBatchSize: s.cfg.MaxBatchSize,
		SafeDelay:    time.Duration(s.cfg.SafeDelayMin) * time.Minute,
		MoveInterval: time.Duration(s.cfg.MoveIntervalMs) * time.Millisecond,
		DryRun:       s.cfg.DryRun,
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
			interval := time.Duration(s.cfg.PollIntervalMin) * time.Minute
			if interval <= 0 {
				interval = 60 * time.Minute
			}
			s.ticker.Reset(interval)
		}
		s.mu.Unlock()
	}()

	s.logger.Info("开始全量整理任务")

	// 检查磁盘空间
	status, err := monitor.CheckDiskSpace(s.cfg.PSTRootPath)
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

// GetState 获取当前状态（供托盘 UI 查询）
func (s *Scheduler) GetState() SchedulerState {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.state
}
