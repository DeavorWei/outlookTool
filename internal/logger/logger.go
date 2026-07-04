package logger

import (
	"os"
	"path/filepath"
	"strings"

	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
	"gopkg.in/natefinch/lumberjack.v2"

	"outlook-archiver/internal/config"
)

// CurrentLogDir 保存当前实际使用的日志目录
var CurrentLogDir string

// LogBroadcast 广播实时日志的通道
var LogBroadcast = make(chan string, 200)

type broadcastSyncer struct{}

func (b *broadcastSyncer) Write(p []byte) (n int, err error) {
	// 非阻塞写入，防止影响正常业务
	select {
	case LogBroadcast <- string(p):
	default:
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
		Filename:   filepath.Join(logDir, "archiver.log"),
		MaxSize:    50,                   // MB
		MaxAge:     cfg.LogRetentionDays, // 天
		LocalTime:  true,
		Compress:   false,
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
	return zap.New(core, zap.AddCaller()), nil
}

func exeDir() string {
	exePath, err := os.Executable()
	if err != nil {
		cwd, _ := os.Getwd()
		return cwd
	}
	dir := filepath.Dir(exePath)
	// 判断是否在 go run 或临时编译目录下运行
	if strings.Contains(dir, os.TempDir()) || strings.Contains(dir, "Temp") || strings.Contains(dir, "tmp") {
		cwd, err := os.Getwd()
		if err == nil {
			return cwd
		}
	}
	return dir
}
