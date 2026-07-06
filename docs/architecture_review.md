# Outlook Auto-Archiver 架构与业务逻辑审查报告

对项目的整体架构和核心业务逻辑进行了详细核查，项目整体设计（尤其是单线程 COM 桥接、Restrict 过滤、按季度路由）非常优秀且成熟。以下是针对潜在风险和可优化方向的详细清单。

## 一、 架构与并发设计审查

### 1. 优势
* **COM 线程安全**：使用了 `COMBridge` 结合通道（Channel）的单线程模型处理所有 `IDispatch` 调用。这是 Go 操作 Outlook COM 的最佳实践，完美避免了 STA（Single-Threaded Apartment）跨线程调用的崩溃问题。
* **资源释放严谨**：代码中大量且正确地使用了 `v.Clear()` 和 `comutil.SafeRelease()`，以及 `defer` 机制，这对于长时间运行的后台守护进程至关重要，有效防止了内存泄漏和 GDI 句柄耗尽。

### 2. 潜在并发数据竞争 (Data Race)
* **配置热重载问题**：在 `internal/scheduler/scheduler.go` 的 `ReloadConfig` 方法中，使用了 `*s.cfg = *newCfg` 来就地更新配置。虽然有 `s.mu.Lock()` 保护，但这只能保证 `scheduler` 内部不冲突。如果此时 `archiver.Archive()` 正在另一个 goroutine 中执行，并在没有加锁的情况下读取 `cfg` 的字段，就会发生并发数据竞争（Data Race），严重时可能导致 Panic。
  * **优化建议**：将配置对象改为 `atomic.Value` 存储，或者在每次触发任务时，将当前的配置深拷贝一份（按值传递）给 `Archiver` 执行，实现真正的无锁并发安全。

## 二、 核心业务逻辑审查

### 1. 邮件遍历与索引性能 (O(N^2) 问题风险)
* **发现**：在 `internal/archiver/restore.go` 中，为了防止由于移动/删除导致的索引错乱，使用了倒序遍历 `for i := count; i >= 1; i--` 来获取邮件 (`items.Item(i)`)。
* **风险**：在 Outlook COM 中，对于极大型的文件夹（如超过 5 万封邮件），通过索引 `Item(i)` 随机访问的时间复杂度在某些版本中是 O(N)。这会导致整个遍历过程变成 O(N^2)，使得归档后期变得极其缓慢。
* **优化建议**：对于纯读取，优先使用 `GetFirst()` 和 `GetNext()`。如果必须移动或删除（会改变集合大小），可以先用 `GetFirst/GetNext` 遍历一遍，收集所有符合条件的邮件的 `EntryID`，存入一个切片（Slice）中。然后遍历这个切片，通过 `GetItemFromID` 获取邮件对象再执行移动/删除。

### 2. 邮件去重策略的鲁棒性
* **发现**：在 `restore.go` 和 `reorganizer.go` 中，识别重复邮件的逻辑是依赖 `Subject` (主题) + `MailTime` (时间戳)。
* **风险**：这种组合有概率产生误判（False Positive）。例如，某些系统告警邮件、Cron 任务通知等，可能会在同一秒内发送多封主题完全相同的邮件。按现有逻辑，它们会被误认为重复邮件而被删除。
* **优化建议**：Outlook 邮件有一个全球唯一的标识符字段 `InternetMessageId`（即邮件头中的 Message-ID）。建议通过 `PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")` 获取此 ID 作为去重的主键 (Key)，这样可以保证 100% 的准确率。

### 3. 系统文件夹识别逻辑 (多语言与定制化问题)
* **发现**：在 `internal/outlook/folder.go` 中，`isSystemReserved` 通过字符串硬编码匹配（如 `"已删除邮件"`, `"deleted items"` 等）来排除系统文件夹。
* **风险**：如果用户的 Windows/Office 是其他语言（如日文、法文），或者企业 Exchange 管理员自定义了系统文件夹名称，硬编码匹配就会失效，导致系统文件夹被错误归档。
* **优化建议**：不要依赖字符串名称。可以通过 `Namespace.GetDefaultFolder()` 获取各个系统默认文件夹（如 Inbox(6), DeletedItems(3), Drafts(16)）的 `EntryID`。在遍历文件夹时，比对当前文件夹的 `EntryID` 与默认系统文件夹的 `EntryID` 是否一致，以此来判断是否为系统保留文件夹。

### 4. PST 文件容量限制与生命周期
* **发现**：程序按季度 (`router.go`) 自动创建 PST 文件。
* **风险**：Outlook 对单个 PST 文件有物理大小限制（通常 Unicode 格式默认上限为 50GB，但建议在 20GB 以内以防损坏）。如果某个季度的邮件量极大（包含大量超大附件），直接写入可能会导致 PST 损坏，且目前代码没有容量检测机制。
* **优化建议**：在 `EnsurePSTFolder` 或移动邮件前，可以通过读取 PST 文件在磁盘上的物理大小 (`os.Stat`)。如果超过设定的安全阈值（如 20GB），则自动触发分卷机制（例如切分为 `2024_Q1_Part2.pst`）。

### 5. 跨 Store 移动的重试机制
* **风险**：由于网络波动（如果是网络磁盘上的 PST）或 Outlook 索引卡顿，`MoveItem` 偶尔会抛出 COM 错误。目前代码是记录日志并增加 `failed` 计数。
* **优化建议**：对于偶发的 COM 错误，可以在捕获到错误后，进行短时间（如 1 秒）的休眠并自动重试 1-2 次，这能极大提高无人值守守护进程的成功率。

## 三、 总结
整体而言，这是一个质量很高的 Go 桌面端守护进程。修复**配置读写的并发安全问题**，并将**系统文件夹判断改为 EntryID 比对**，是短期内性价比最高的优化项。
