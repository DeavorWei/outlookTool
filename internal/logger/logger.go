package logger

import (
	"os"
	"path/filepath"
	"strings"
	"sync/atomic"
	"time"

	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
	"gopkg.in/natefinch/lumberjack.v2"

	"outlook-archiver/internal/config"
)

// CurrentLogDir 保存当前实际使用的日志目录
var CurrentLogDir string

// LogBroadcast 广播实时日志的通道
var LogBroadcast = make(chan string, 1000)
var dropCount int32

type broadcastSyncer struct{}

func (b *broadcastSyncer) Write(p []byte) (n int, err error) {
	// 非阻塞写入，防止影响正常业务
	select {
	case LogBroadcast <- string(p):
	default:
		atomic.AddInt32(&dropCount, 1)
	}
	return len(p), nil
}

func (b *broadcastSyncer) Sync() error { return nil }

// InitLogger 初始化 zap + lumberjack 日志模块
func InitLogger(cfg *config.Config) (*zap.Logger, error) {
	logDir := filepath.Join(exeDir(), "logs")
	if err := os.MkdirAll(logDir, 0755); err != nil {
		logDir = filepath.Join(os.TempDir(), "OutlookAutoArchiver", "logs")
		if err2 := os.MkdirAll(logDir, 0755); err2 != nil {
			return nil, err
		}
	}
	CurrentLogDir = logDir

	w := &lumberjack.Logger{
		Filename:  filepath.Join(logDir, "archiver.log"),
		MaxSize:   50,                   // MB
		MaxAge:    cfg.LogRetentionDays, // 天
		LocalTime: true,
		Compress:  false,
	}

	encoderConfig := zap.NewProductionEncoderConfig()
	encoderConfig.EncodeTime = zapcore.ISO8601TimeEncoder

	fileLevel := zapcore.InfoLevel
	if cfg != nil && cfg.DebugLog {
		fileLevel = zapcore.DebugLevel
	}

	fileCore := zapcore.NewCore(
		zapcore.NewJSONEncoder(encoderConfig),
		zapcore.AddSync(w),
		fileLevel,
	)

	// 同时输出到控制台（方便开发调试）
	consoleEncoderConfig := zap.NewDevelopmentEncoderConfig()
	consoleEncoderConfig.EncodeTime = zapcore.ISO8601TimeEncoder
	consoleCore := zapcore.NewCore(
		zapcore.NewConsoleEncoder(consoleEncoderConfig),
		zapcore.AddSync(os.Stdout),
		zapcore.DebugLevel,
	)

	// Web SSE 广播 Core，使用 JSON 格式以便于前端解析颜色和字段
	broadcastCore := zapcore.NewCore(
		zapcore.NewJSONEncoder(encoderConfig),
		zapcore.AddSync(&broadcastSyncer{}),
		zapcore.DebugLevel,
	)

	core := zapcore.NewTee(fileCore, consoleCore, broadcastCore)
	logger := zap.New(core, zap.AddCaller())
	// M8：周期上报被丢弃的实时日志条数，避免静默丢日志
	go reportDroppedLogs(logger)
	return logger, nil
}

// reportDroppedLogs 周期上报被丢弃的实时日志条数（SSE 广播通道溢出时会发生）
func reportDroppedLogs(logger *zap.Logger) {
	ticker := time.NewTicker(time.Minute)
	defer ticker.Stop()
	for range ticker.C {
		dropped := atomic.SwapInt32(&dropCount, 0)
		if dropped > 0 && logger != nil {
			logger.Warn("日志广播通道溢出，部分实时日志已被丢弃", zap.Int32("dropped_count", dropped))
		}
	}
}

func exeDir() string {
	exePath, err := os.Executable()
	if err != nil {
		cwd, _ := os.Getwd()
		return cwd
	}
	dir := filepath.Dir(exePath)
	dirLower := strings.ToLower(dir)
	// 仅依据系统临时目录前缀判定（避免 "temp"/"tmp" 子串误判，如 "mytemplate" 目录）
	if strings.Contains(dirLower, strings.ToLower(os.TempDir())) {
		cwd, err := os.Getwd()
		if err == nil {
			return cwd
		}
	}
	return dir
}
