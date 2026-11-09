# QuickExcel

高性能 .xlsx 流式导入导出：DataAnnotations 验证、图片、模板填充。

从原 QuickNet 单仓拆分出的独立项目。MIT 协议，零第三方依赖。

## 安装

```bash
dotnet add package QuickExcel --prerelease
```

## 快速上手

```csharp
using QuickExcel;

// 导出
Excel.SaveAs("orders.xlsx", orderList, new ExportOptions { BoldHeader = true, FreezeHeaderRow = true });

// 导入 + DataAnnotations 验证（错误合并不短路，行号/列名/单元格引用逐条定位）
ImportResult<Employee> result = Excel.Import<Employee>("in.xlsx");
if (result.HasErrors)
    foreach (var e in result.Errors)
        Console.WriteLine($"第{e.RowNumber}行 {e.CellReference} [{e.Header}] {e.Message}");

// 模板填充：{{标量}} / {{items.Field}} 集合行展开 / {{@img}} 图片占位
Excel.SaveAsByTemplate("report.xlsx", "template.xlsx", new Dictionary<string, object?>
{
    ["Title"] = "季度报告",
    ["items"] = orderList,
    ["logo"]  = ExcelImage.FromFile("logo.png"),
});
```

## 特性

- SAX 式写出 + 前向读取，内存与行数无关
- 图片列（`ExcelImage`）导出为浮动图片、导入按锚点回填，兼容 WPS `DISPIMG`；相同字节自动去重
- 合并单元格导入自动回填（`FillMergedCells`）；多数据源（POCO / 字典 / ExpandoObject / DataTable / 多 Sheet）

## 构建与测试

```bash
dotnet build QuickExcel.slnx
dotnet test tests/QuickExcel.Tests/QuickExcel.Tests.csproj
```

## 文档

见 [docs/quickexcel-usage.md](docs/quickexcel-usage.md)。

## License

MIT
