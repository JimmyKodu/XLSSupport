# XLSSupport

NPOI 2.7.5 Excel文件兼容性测试（.NET 9）

## 项目说明

本项目演示了 NPOI 2.7.5 在 .NET 9 环境下写入 Excel 文件，并与多个主流 Excel 库的读写兼容性测试。

不要相信网上说 NPOI 与其他库不兼容的言论！本项目用实际代码证明了完全兼容！

## 测试环境

- **.NET 版本**: .NET 9.0
- **NPOI 版本**: 2.7.5

## 测试的库及版本

| 库名称 | 版本 | XLS 支持 | XLSX 支持 | 测试结果 |
|--------|------|----------|-----------|----------|
| NPOI | 2.7.5 | ✓ 读写 | ✓ 读写 | ✅ 通过 |
| MiniExcel | 1.41.4 | ✗ | ✓ 读 | ✅ 通过 |
| ExcelDataReader | 3.8.0 | ✓ 读 | ✓ 读 | ✅ 通过 |
| DocumentFormat.OpenXml | 3.3.0 | ✗ | ✓ 读 | ✅ 通过 |
| ClosedXML | 0.105.0 | ✗ | ✓ 读 | ✅ 通过 |
| EPPlus | 4.5.3.3 | ✗ | ✓ 读 | ✅ 通过 |

## 重要发现

### XLS 格式（二进制格式）
- ✅ **NPOI 2.7.5 (HSSF)** 可以完美读写 XLS 文件
- ✅ **ExcelDataReader 3.8.0** 可以读取 NPOI 创建的 XLS 文件
- ❌ MiniExcel、ClosedXML、EPPlus、DocumentFormat.OpenXml 不支持 XLS 格式

### XLSX 格式（OpenXML 格式）
- ✅ **所有测试的库都能成功读取** NPOI 2.7.5 创建的 XLSX 文件
- ✅ 支持中文字符（如"张三"、"李四"、"王五"）
- ✅ 支持各种数据类型：整数、字符串、浮点数、布尔值

## 如何运行

### 前置条件
- 安装 .NET 9.0 SDK

### 运行步骤

```bash
cd XLSCompatibilityTest
dotnet restore
dotnet build
dotnet run
```

## 测试数据

程序会创建包含以下测试数据的 Excel 文件：

| ID | Name | Age | Score | Active |
|----|------|-----|-------|--------|
| 1 | 张三 | 25 | 95.5 | true |
| 2 | 李四 | 30 | 87.3 | false |
| 3 | 王五 | 28 | 92.1 | true |
| 4 | John Doe | 35 | 88.8 | true |
| 5 | Jane Smith | 29 | 96.7 | false |

## 测试流程

1. **写入测试**: 使用 NPOI 2.7.5 创建 XLS 和 XLSX 文件
2. **读取测试**: 使用各个库读取创建的文件
3. **验证**: 检查读取的数据是否正确

## 结论

✅ **NPOI 2.7.5 在 .NET 9 上完全兼容！**

- NPOI 写入的 XLS 文件可以被 ExcelDataReader 正确读取
- NPOI 写入的 XLSX 文件可以被所有主流 Excel 库正确读取
- 支持中英文混合、多种数据类型
- 在生产环境中可以放心使用

## 项目结构

```
XLSSupport/
├── README.md
└── XLSCompatibilityTest/
    ├── Program.cs                    # 主程序，包含所有测试代码
    └── XLSCompatibilityTest.csproj  # 项目文件
```

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！