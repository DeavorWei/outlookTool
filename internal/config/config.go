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
			err = saveDefaultConfigWithComments(path, config)
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

// saveDefaultConfigWithComments 保存带有详细注释的默认配置文件
func saveDefaultConfigWithComments(path string, config *Config) error {
	defaultYAML := fmt.Sprintf(`# Outlook Auto-Archiver 配置文件
# ==========================================
# 该配置文件在程序首次运行时自动生成。修改配置后，请重启程序或通过托盘菜单重新加载以使配置生效。
# ==========================================

# pst_root_path: PST 文件的存储根目录
# 作用: 所有归档生成的 PST 文件（按季度命名，如 2024_Q2.pst）都会存放在该目录下。
# 注意: 请确保该目录所在的磁盘有充足的空间（建议预留 5GB 以上）。路径分隔符在 Windows 下使用双反斜杠 \\ 或单斜杠 / 均可。
# 示例: "D:\\OutlookArchives" 或 "C:\\Users\\admin\\Documents\\Outlook Files"
pst_root_path: "%s"

# poll_interval_minutes: 轮询间隔（分钟）
# 作用: 后台进程每隔多久自动扫描一次邮箱中的邮件。
# 示例: 10 (表示每 10 分钟执行一次归档检查)
poll_interval_minutes: 10

# safe_delay_minutes: 安全延迟时间（分钟）
# 作用: 仅处理早于“当前时间减去安全延迟时间”的邮件。这是为了防止处理正在与 Exchange 服务器同步的邮件，避免引发 RPC 锁死。
# 示例: 10 (表示现在是 10:30 的话，只处理 10:20 之前的邮件)
safe_delay_minutes: 10

# max_batch_size: 单次最大处理邮件数
# 作用: 每次轮询任务中，最多移动多少封邮件。该额度由所有文件夹共享。限制批次可以防止瞬间资源占用过高导致 Outlook 假死。
# 示例: 500 (单次最多移动 500 封邮件，超过部分留到下一次轮询处理)
max_batch_size: 500

# archive_mode: 归档模式
# 作用: 决定哪些文件夹会被归档。
# 可选值:
#   - "all" : 归档所有文件夹（包含默认文件夹和用户自定义文件夹），除了 exclude_folders 中列出的文件夹。
#   - "list": 仅归档 include_folders 中明确列出的文件夹，忽略其他文件夹。
# 示例: "all"
archive_mode: all

# exclude_folders: 排除归档的文件夹列表
# 作用: 当 archive_mode 为 "all" 时，这些列出的文件夹将不会被归档。
# 注意: 系统保留文件夹（如 已删除邮件、垃圾邮件、草稿、发件箱、同步问题等）在代码层面已默认排除，无需在此处重复添加。
# 示例: ["RSS 订阅", "对话历史记录"]
exclude_folders: []

# include_folders: 仅归档的文件夹列表
# 作用: 当 archive_mode 为 "list" 时，只有在此处列出的文件夹才会被归档。
# 注意: 支持多级文件夹，用斜杠分隔。
# 示例: ["Inbox", "SentItems", "项目A", "客户跟进/重要"]
include_folders: []

# log_retention_days: 日志保留天数
# 作用: 运行日志文件最多保留多少天，超过天数的历史日志会被自动清理。
# 示例: 7 (保留最近 7 天的日志)
log_retention_days: 7

# move_interval_ms: 邮件移动间隔（毫秒）
# 作用: 成功移动一封邮件后，程序强制休眠的时间。用于给予 Outlook 和 MAPI 引擎喘息时间，防止界面卡顿。
# 示例: 50 (每移动一封邮件休眠 50 毫秒)
move_interval_ms: 50

# dry_run: 模拟运行模式
# 作用: 开启后，程序会完整执行归档的扫描、过滤和计算逻辑，并在日志中打印“将会移动”的详情，但**不会真正执行物理移动**。
# 示例: false (关闭模拟运行，执行真实的归档操作)
dry_run: false

# copy_only: 仅复制模式
# 作用: 开启后，邮件只会被复制到 PST 中，而不会从源邮箱中删除。
# 示例: false
copy_only: false

# debug_log: 调试日志开关
# 作用: 开启后，会将更为详细的 Debug 级别日志输出到文件中，便于排查疑难问题。
# 示例: false (日常运行建议关闭以节省磁盘空间)
debug_log: false

# legacy_pst_scan_paths: 历史/第三方 PST 扫描目录列表
# 作用: 用于“全量整理”功能。程序会扫描这些目录下的所有 .pst 文件，并将其中被错误归档（时间错乱）的邮件统一纠正到正确的季度 PST 中。
# 示例: ["D:\\OutlookArchives", "C:\\Users\\admin\\Documents\\Outlook Files"]
legacy_pst_scan_paths: []
`, filepath.ToSlash(config.PSTRootPath))

	return os.WriteFile(path, []byte(defaultYAML), 0644)
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
