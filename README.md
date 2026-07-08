<div align="center">

# 📬 Outlook Auto-Archiver

**Windows 平台 Outlook 邮箱自动归档守护进程**

[![Go Version](https://img.shields.io/badge/Go-1.26-blue?logo=go)](https://go.dev/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011-blue?logo=windows)](https://www.microsoft.com/windows/)
[![License](https://img.shields.io/badge/License-MIT-green)](#license)

[功能特性](#-功能特性) • [快速开始](#-快速开始) • [配置说明](#-配置说明) • [项目架构](#-项目架构) • [📖 用户指南](docs/USER_GUIDE.md)

</div>

---

## 📖 项目简介

**Outlook Auto-Archiver** 是一个轻量级的 Windows 后台守护程序，通过 COM 接口与 Microsoft Outlook 无缝交互，自动将 Exchange 邮箱中的邮件按季度归档到本地 PST 文件中。

### 🎯 解决的核心痛点

| 痛点 | 解决方案 |
|------|----------|
| 📭 邮箱容量受限（如 1GB 上限）导致爆仓 | 自动将历史邮件迁移至本地 PST，释放服务器空间 |
| 📚 历史邮件堆积无法有效管理 | 按季度自动分类，严格基于邮件真实时间归档 |
| 🔀 第三方归档插件分类错乱 | 全量整理功能可自动纠偏已有 PST 中的错误分类 |

程序运行于后台，对用户完全无感知，严格按照邮件的**实际发送/接收时间**进行归档路由，确保任何时候归档都能分发到正确的季度 PST 文件中。

---

## ✨ 功能特性

### 🤖 核心归档能力

- **自动定时归档** — 按可配置的轮询间隔（默认 10 分钟）静默扫描邮箱，将超过保留天数的邮件按季度归档到本地 PST 文件
- **严格时间基准** — 使用邮件的真实发送/接收时间而非系统当前时间，确保归档分类准确无误
- **镜像文件夹结构** — PST 中的目录层级完美还原源邮箱结构，支持自定义多级文件夹

### 🔧 高级操作

- **全量整理** — 一键执行深度扫描：归档所有邮件 + 纠偏已有 PST 中的错误分类 + 迁移第三方 PST 文件
- **邮件还原** — 将 PST 中的邮件还原回 Exchange 邮箱，支持去重和自动清理空 PST
- **模拟运行模式** — Dry Run 模式可预览归档操作而不实际移动邮件，适合首次部署验证

### 🖥️ 用户界面

- **系统托盘控制** — 右键菜单快速操作：立即执行、全量整理、还原、配置管理等
- **Web 控制台** — 嵌入式 Web 界面，提供配置中心、实时日志流（SSE）、操作面板

### 🛡️ 安全与稳定性

- **磁盘空间监控** — 自动检测磁盘剩余空间和 PST 文件大小，空间不足时告警并暂停操作
- **防多开保护** — 内置系统级单例锁，防止程序重复运行
- **开机自启** — 支持通过注册表配置 Windows 开机自启动
- **热重载配置** — 修改配置后无需重启程序即可生效

---

## ⚙️ 系统要求

| 项目 | 要求 |
|------|------|
| **操作系统** | Windows 10 / Windows 11 |
| **邮件客户端** | Microsoft Office LTSC 2024 专业增强版 或 经典版 Classic Outlook |
| **磁盘空间** | PST 存储目录建议预留至少 5GB 可用空间 |
| **运行权限** | 当前普通用户权限（User Session），**禁止部署为 Windows 服务（Session 0）** |

> ⚠️ **重要**：本程序**不支持**基于 WebView2 的"新版 Outlook"，仅支持传统 COM 接口的 Classic Outlook。

---

## 🚀 快速开始

### 1. 获取程序

从 [Releases](https://github.com/deavorwei/outlookTool/releases) 页面下载预编译的 `outlook-archiver.exe`，或从源码构建：

```powershell
# 克隆仓库
git clone https://github.com/deavorwei/outlookTool.git
cd outlookTool

# 编译
go build -o outlook-archiver.exe .
```

### 2. 部署运行

```powershell
# 将编译好的可执行文件放置到合适目录（建议空间充足的磁盘）
# 直接双击运行，或在命令行中执行：
.\outlook-archiver.exe
```

程序启动后将在系统托盘区显示图标，并自动在同目录下生成：
- `config.yaml` — 配置文件
- `logs/` — 日志目录
- `OutlookArchives/` — 默认 PST 存储目录

### 3. 首次验证（推荐）

首次部署时建议先以模拟模式运行，确认归档计划符合预期：

```yaml
# 编辑 config.yaml
dry_run: true
```

此时程序会执行完整的分析逻辑，在日志中打印"将会移动"的详情，但不实际移动邮件。确认无误后将 `dry_run` 设置为 `false` 即可正式运行。

---

## 📝 配置说明

配置文件 `config.yaml` 在程序首次运行时自动生成，包含详细的中文注释。

### 完整配置项一览

| 配置项                 | 类型       | 默认值              | 说明                                                    |
| ---------------------- | ---------- | ------------------- | ------------------------------------------------------- |
| `pst_root_path`        | `string`   | `./OutlookArchives` | PST 文件存储根目录                                       |
| `poll_interval_minutes`| `int`      | `10`                | 自动扫描轮询间隔（分钟）                                 |
| `safe_delay_minutes`   | `int`      | `10`                | 安全延迟，防止与服务器同步冲突                            |
| `retain_days`          | `int`      | `30`                | 保留近期邮件不归档的天数（0 = 不保留）                    |
| `max_batch_size`       | `int`      | `500`               | 单次最大处理邮件数量                                     |
| `archive_mode`         | `string`   | `"all"`             | `all`（归档所有，排除黑名单）/ `list`（仅白名单）         |
| `exclude_folders`      | `[]string` | `[]`                | 排除的文件夹列表（黑名单模式）                            |
| `include_folders`      | `[]string` | `[]`                | 包含的文件夹列表（白名单模式）                            |
| `log_retention_days`   | `int`      | `7`                 | 日志文件保留天数                                         |
| `move_interval_ms`     | `int`      | `50`                | 移动单封邮件后的休眠时间（毫秒），避免 Outlook 卡顿       |
| `dry_run`              | `bool`     | `false`             | 模拟运行模式，仅打印日志不实际移动                        |
| `copy_only`            | `bool`     | `false`             | 仅复制模式，邮件只复制到 PST，源邮件不删除                 |
| `debug_log`            | `bool`     | `false`             | 开启 Debug 级别详细日志                                  |
| `legacy_pst_scan_paths`| `[]string` | `[]`                | 第三方 PST 文件扫描路径（用于全量整理迁移）               |
| `include_mounted_psts` | `bool`     | `true`              | 是否自动扫描 Outlook 中已挂载的 PST 文件                  |
| `stream_block_size`    | `int`      | `1000`              | 流式遍历邮件时的分块大小                                  |

> 💡 **提示**：修改配置文件后，可通过系统托盘菜单的"重新加载配置"选项使新配置生效，无需重启程序。

---

## 📋 使用说明

### 🖱️ 系统托盘菜单

程序在后台运行时，右键点击系统托盘图标可进行以下操作：

| 菜单项 | 说明 |
|--------|------|
| **立即执行一次** | 手动触发一次归档任务（受最大批次限制） |
| **全量整理** | 暂停日常任务，执行深度扫描与纠偏 |
| **邮件还原** | 将 PST 中的邮件还原回 Exchange 邮箱 |
| **打开 Web 控制台** | 在浏览器中打开嵌入式 Web 管理界面 |
| **打开日志目录** | 快速查看程序运行日志 |
| **打开配置文件** | 调用系统默认编辑器打开 `config.yaml` |
| **重新加载配置** | 热重载配置，无需重启程序 |
| **开机自启** | 开启/关闭随系统自动启动 |
| **退出** | 优雅关闭程序，释放 COM 资源 |

### 🌐 Web 控制台

程序启动后会运行一个嵌入式 Web 服务器，提供以下功能：

- **配置中心** — 在线查看和修改配置项
- **实时日志** — 通过 SSE（Server-Sent Events）实时推送运行日志
- **操作面板** — 执行归档、全量整理、还原等操作

---

## 🏗️ 项目架构

```
outlookTool/
├── main.go                          # 程序入口
├── config.yaml                      # 配置文件（首次运行自动生成）
├── internal/
│   ├── archiver/                    # 📦 归档引擎
│   │   ├── archiver.go              #    常规归档逻辑
│   │   ├── reorganizer.go           #    全量整理 + PST 纠偏
│   │   └── restore.go               #    邮件还原
│   ├── config/                      # ⚙️ 配置管理（YAML 读写、校验）
│   ├── logger/                      # 📝 日志系统（zap + lumberjack + SSE 广播）
│   ├── monitor/                     # 📊 磁盘/PST 大小监控
│   ├── mutex/                       # 🔒 系统级单例锁（防重复运行）
│   ├── outlook/                     # 🔌 Outlook COM 桥接层
│   │   ├── bridge.go                #    COM 线程管理与进程控制
│   │   ├── folder.go                #    邮箱文件夹遍历
│   │   ├── mail.go                  #    邮件操作（移动/复制/删除）
│   │   └── pst.go                   #    PST 文件管理
│   ├── registry/                    # 📋 注册表操作（开机自启）
│   ├── router/                      # 🧭 季度路由逻辑
│   ├── scheduler/                   # ⏰ 定时任务调度（状态机）
│   ├── server/                      # 🌐 Web 服务器 + SSE 日志流
│   │   └── web/                     #    嵌入式前端页面
│   └── tray/                        # 🖥️ 系统托盘 UI
└── docs/                            # 📚 项目文档
    └── USER_GUIDE.md                #    用户指南
```

### 模块职责

| 模块 | 职责 |
|------|------|
| `archiver` | 核心归档引擎，负责邮件的季度归档、全量整理和还原操作 |
| `config` | YAML 配置文件的读取、写入和校验 |
| `logger` | 基于 zap 的高性能结构化日志，支持文件轮转和 SSE 实时广播 |
| `monitor` | 磁盘空间和 PST 文件大小的监控与告警 |
| `mutex` | 系统级单例锁，确保同一时间只有一个实例运行 |
| `outlook` | Outlook COM 接口的 Go 语言封装层 |
| `registry` | Windows 注册表操作，实现开机自启功能 |
| `router` | 根据邮件时间计算目标季度 PST 文件的路由逻辑 |
| `scheduler` | 定时任务调度器，管理归档任务的状态机 |
| `server` | 嵌入式 Web 服务器，提供 REST API 和 SSE 日志流 |
| `tray` | 系统托盘 UI 和右键菜单交互 |

---

## 🔨 从源码构建

### 前置条件

- [Go 1.26+](https://go.dev/dl/)
- Windows 10/11 + Classic Outlook 已安装
- CGO 工具链（用于编译系统托盘依赖）

### 构建命令

```powershell
# 下载依赖
go mod download

# 开发构建
go build -o outlook-archiver.exe .

# 生产构建（优化体积）
go build -ldflags="-s -w" -o outlook-archiver.exe .
```

---

## 📚 技术栈

| 技术 | 用途 |
|------|------|
| **Go 1.26** | 主要编程语言 |
| **Windows COM/OLE Automation** | 与 Microsoft Outlook 交互的核心接口 |
| **系统托盘（systray）** | 原生 Windows 托盘图标和菜单 |
| **嵌入式 Web 控制台** | 内嵌 HTML/JS 前端，提供 Web 管理界面 |
| **SSE (Server-Sent Events)** | 实时日志推送 |
| **Zap + Lumberjack** | 高性能结构化日志与文件轮转 |

---

## 📦 依赖列表

| 依赖 | 用途 |
|------|------|
| [`github.com/getlantern/systray`](https://github.com/getlantern/systray) | 跨平台系统托盘支持 |
| [`github.com/go-ole/go-ole`](https://github.com/go-ole/go-ole) | Windows COM/OLE Automation Go 绑定 |
| [`go.uber.org/zap`](https://github.com/uber-go/zap) | 高性能结构化日志库 |
| [`golang.org/x/sys`](https://pkg.go.dev/golang.org/x/sys) | Windows API 调用（进程、磁盘、注册表） |
| [`gopkg.in/natefinsh/lumberjack.v2`](https://github.com/natefinch/lumberjack) | 日志文件自动轮转 |
| [`gopkg.in/yaml.v3`](https://gopkg.in/yaml.v3) | YAML 配置文件解析 |

---

## ❓ 常见问题

### Q: 程序启动后没有任何反应？
A: 程序仅在后台运行，请查看系统托盘区域（任务栏右下角）是否有程序图标。首次运行时需要 Outlook 已经启动。

### Q: 归档后邮件从 Outlook 中消失了？
A: 这是正常行为。邮件被移动到了本地 PST 文件中。您可以在 Outlook 中通过"文件 > 打开和导出 > 打开 Outlook 数据文件"来查看归档的 PST。

### Q: 如何恢复被错误归档的邮件？
A: 使用系统托盘菜单中的"邮件还原"功能，程序会将 PST 中的邮件自动还原回 Exchange 邮箱，并支持去重。

### Q: 配置修改后没有生效？
A: 修改 `config.yaml` 后，需要通过托盘菜单的"重新加载配置"或重启程序使配置生效。

### Q: 磁盘空间不足会怎样？
A: 程序会自动监控磁盘空间。剩余空间不足 1GB 时发出警告，不足 500MB 时暂停所有归档操作。

---

> 📖 **更多详细说明请参阅 [用户指南](docs/USER_GUIDE.md)**，包含完整的使用教程、高级配置说明和故障排查指南。

---

## 📄 License

本项目采用 [MIT License](LICENSE) 开源协议。

---

<div align="center">

**[⬆ 回到顶部](#-outlook-auto-archiver)**

</div>
