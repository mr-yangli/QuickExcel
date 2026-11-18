# 更新日志

QuickExcel 各版本的变更，倒序排列（最新在上）。

> **口径**
> - 版本号跟 .NET 主版本走（.NET 10 → 10.x）。破坏性改动升主版本号，新增升次版本号（10.0.0 → 10.1.0），缺陷修复升修订号。
> - 「未发布」收 main 已合入、尚未定版的改动；定版时改节名为版本号并标「待发布」（`Directory.Build.props` 的 `<Version>` 同步就位），上包那天把「待发布」换成日期。
> - 分档为 新增 / 变更 / 修复 / 移除；标 **破坏性** 的改动，消费方升级时要动代码。
> - 版本日期取该版本在 NuGet 上可还原之日，不是提交日。

## 未发布

暂无。

## 10.1.0 — 2026-11-18

### 新增

- **多目标 `net10.0;net11.0`**：`QuickExcel` 同包双 TFM（单包双资产），消费方按宿主框架自动还原对应资产，`net11.0` 消费方不用再退回 `net10.0` 资产。包描述与 tags 同步标 .NET 10/11。属新增（非破坏），按家规升次版本号。

## 10.0.0 — 2026-11-09

首个公开版本：QuickExcel 从内部单仓拆为独立仓库（GitHub `mr-yangli/QuickExcel`、Gitee `mr_yangli/quick-excel`），本族 1 颗包 `QuickExcel` 上 NuGet。零第三方依赖，MIT。下面列的是发布内容面，不是相对上一版的 diff——10.0.0 之前没有公开版本。

### 新增

- **导出**：`Excel.SaveAs` 走 SAX 式流式写出，内存与行数无关；`ExportOptions` 支持表头加粗、冻结表头行等版式项。
- **导入与验证**：`Excel.Import<T>` 前向读取 + DataAnnotations 校验；错误合并不短路，逐条给出行号、列名、单元格引用（`ImportResult<T>.HasErrors` / `Errors`）。
- **图片**：`ExcelImage` 导出为浮动图片，导入按锚点回填；兼容 WPS 的 `DISPIMG`；相同字节自动去重，不重复入包。
- **合并单元格**：导入时按 `FillMergedCells` 回填到每一格，避免只有左上角有值。
- **多数据源**：POCO / 字典 / `ExpandoObject` / `DataTable` / 多 Sheet 均可作为一次写出的来源。
- **模板填充**：`Excel.SaveAsByTemplate` 支持 `{{标量}}`、`{{items.Field}}` 集合行展开、`{{@img}}` 图片占位。
- **测试面**：`tests/QuickExcel.Tests` —— 导出/导入往返（`RoundtripTests`）、验证错误定位（`ValidationTests`）、图片与字节去重（`ImageTests`）、WPS `DISPIMG`（`DispImgTests`）、合并单元格（`MergeCellTests`）、模板展开（`TemplateTests`）、转义闸门（`EscapingGuardTests`）、样式（`StyleTests`）。
