package config

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"gopkg.in/yaml.v3"
)

type Config struct {
	PSTRootPath        string   `yaml:"pst_root_path"`
	PollIntervalMin    int      `yaml:"poll_interval_minutes"`
	SafeDelayMin       int      `yaml:"safe_delay_minutes"`
	MaxBatchSize       int      `yaml:"max_batch_size"`
	ArchiveMode        string   `yaml:"archive_mode"` // "all" | "list"
	ExcludeFolders     []string `yaml:"exclude_folders"`
	IncludeFolders     []string `yaml:"include_folders"`
	LogRetentionDays   int      `yaml:"log_retention_days"`
	MoveIntervalMs     int      `yaml:"move_interval_ms"`
	DryRun             bool     `yaml:"dry_run"` // true = 仅日志，不执行 Move
	CopyOnly           bool     `yaml:"copy_only"` // true = 仅复制，不执行 Move
	DebugLog           bool     `yaml:"debug_log"` // true = 将 Debug 级别日志输出到文件
	LegacyPSTScanPaths []string `yaml:"legacy_pst_scan_paths"`
}

func DefaultConfig() *Config {
	exePath, err := os.Executable()
	if err != nil {
		exePath = "."
	}
	defaultRoot := filepath.Join(filepath.Dir(exePath), "OutlookArchives")

	return &Config{
		PSTRootPath:        defaultRoot,
		PollIntervalMin:    10,
		SafeDelayMin:       10,
		MaxBatchSize:       500,
		ArchiveMode:        "all",
		ExcludeFolders:     []string{},
		IncludeFolders:     []string{},
		LogRetentionDays:   7,
		MoveIntervalMs:     50,
		DryRun:             false,
		CopyOnly:           false,
		DebugLog:           false,
		LegacyPSTScanPaths: []string{},
	}
}

// LoadConfig 从指定路径加载配置文件，如果不存在则创建默认配置
func LoadConfig(path string) (*Config, error) {
	config := DefaultConfig()

	data, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			// 如果配置文件不存在，则创建默认配置文件
			_ = os.MkdirAll(config.PSTRootPath, 0755)
			err = SaveConfig(path, config)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Warning: failed to save default config to %s: %v. Using in-memory default.\n", path, err)
			}
			if err := ValidateConfig(config); err != nil {
				return nil, fmt.Errorf("invalid default configuration: %w", err)
			}
			return config, nil
		}
		return nil, fmt.Errorf("failed to read config file: %w", err)
	}

	err = yaml.Unmarshal(data, config)
	if err != nil {
		return nil, fmt.Errorf("failed to parse config file: %w", err)
	}

	// 完整性校验
	if err := ValidateConfig(config); err != nil {
		return nil, fmt.Errorf("invalid configuration: %w", err)
	}

	return config, nil
}

// SaveConfig 保存配置到文件
func SaveConfig(path string, config *Config) error {
	data, err := yaml.Marshal(config)
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0644)
}

// ValidateConfig 校验配置有效性
func ValidateConfig(cfg *Config) error {
	if cfg.PSTRootPath == "" {
		return errors.New("pst_root_path is required")
	}
	if cfg.PollIntervalMin <= 0 {
		return errors.New("poll_interval_minutes must be > 0")
	}
	if cfg.SafeDelayMin < 0 {
		return errors.New("safe_delay_minutes must be >= 0")
	}
	if cfg.MaxBatchSize < 0 {
		return errors.New("max_batch_size must be >= 0")
	}
	if cfg.ArchiveMode != "all" && cfg.ArchiveMode != "list" {
		return errors.New("archive_mode must be 'all' or 'list'")
	}
	if cfg.LogRetentionDays <= 0 {
		return errors.New("log_retention_days must be > 0")
	}
	if cfg.MoveIntervalMs < 0 {
		return errors.New("move_interval_ms must be >= 0")
	}

	info, err := os.Stat(cfg.PSTRootPath)
	if err != nil {
		if os.IsNotExist(err) {
			return fmt.Errorf("pst_root_path '%s' does not exist", cfg.PSTRootPath)
		}
		return fmt.Errorf("failed to stat pst_root_path: %w", err)
	}
	if !info.IsDir() {
		return fmt.Errorf("pst_root_path '%s' is not a directory", cfg.PSTRootPath)
	}

	return nil
}
