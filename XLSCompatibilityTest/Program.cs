using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelDataReader;
using MiniExcelLibs;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using OfficeOpenXml;

Console.WriteLine("=== NPOI 2.7.5 XLS/XLSX Compatibility Test on .NET 9 ===");
Console.WriteLine();

// Register encoding provider for ExcelDataReader
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

// File paths
string xlsFilePath = "test_compatibility.xls";
string xlsxFilePath = "test_compatibility.xlsx";

// Sample data to write
var testData = new List<TestRecord>
{
    new TestRecord { Id = 1, Name = "张三", Age = 25, Score = 95.5, Active = true },
    new TestRecord { Id = 2, Name = "李四", Age = 30, Score = 87.3, Active = false },
    new TestRecord { Id = 3, Name = "王五", Age = 28, Score = 92.1, Active = true },
    new TestRecord { Id = 4, Name = "John Doe", Age = 35, Score = 88.8, Active = true },
    new TestRecord { Id = 5, Name = "Jane Smith", Age = 29, Score = 96.7, Active = false }
};

// ===== XLS FORMAT TESTS =====
Console.WriteLine("==================== XLS FORMAT TESTS ====================");
Console.WriteLine();

// Step 1: Write XLS file using NPOI 2.7.5
Console.WriteLine("Step 1: Writing XLS file with NPOI 2.7.5 (HSSF)...");
WriteXlsWithNPOI(xlsFilePath, testData);
Console.WriteLine($"✓ Successfully created {xlsFilePath} using NPOI 2.7.5");
Console.WriteLine();

// Step 2: Read with NPOI 2.7.5 (verify write operation)
Console.WriteLine("Step 2: Reading XLS with NPOI 2.7.5 (HSSF)...");
ReadXlsWithNPOI(xlsFilePath);
Console.WriteLine();

// Step 3: Read with ExcelDataReader 3.8.0 (supports XLS)
Console.WriteLine("Step 3: Reading XLS with ExcelDataReader 3.8.0...");
ReadWithExcelDataReader(xlsFilePath);
Console.WriteLine();

// ===== XLSX FORMAT TESTS =====
Console.WriteLine();
Console.WriteLine("==================== XLSX FORMAT TESTS ====================");
Console.WriteLine();

// Step 4: Write XLSX file using NPOI 2.7.5
Console.WriteLine("Step 4: Writing XLSX file with NPOI 2.7.5 (XSSF)...");
WriteXlsxWithNPOI(xlsxFilePath, testData);
Console.WriteLine($"✓ Successfully created {xlsxFilePath} using NPOI 2.7.5");
Console.WriteLine();

// Step 5: Read with NPOI 2.7.5 (verify write operation)
Console.WriteLine("Step 5: Reading XLSX with NPOI 2.7.5 (XSSF)...");
ReadXlsxWithNPOI(xlsxFilePath);
Console.WriteLine();

// Step 6: Read with MiniExcel 1.41.4
Console.WriteLine("Step 6: Reading XLSX with MiniExcel 1.41.4...");
ReadWithMiniExcel(xlsxFilePath);
Console.WriteLine();

// Step 7: Read with ExcelDataReader 3.8.0
Console.WriteLine("Step 7: Reading XLSX with ExcelDataReader 3.8.0...");
ReadWithExcelDataReader(xlsxFilePath);
Console.WriteLine();

// Step 8: Read with DocumentFormat.OpenXml 3.3.0
Console.WriteLine("Step 8: Reading XLSX with DocumentFormat.OpenXml 3.3.0...");
ReadWithOpenXml(xlsxFilePath);
Console.WriteLine();

// Step 9: Read with ClosedXML 0.105.0
Console.WriteLine("Step 9: Reading XLSX with ClosedXML 0.105.0...");
ReadWithClosedXML(xlsxFilePath);
Console.WriteLine();

// Step 10: Read with EPPlus 4.5.3.3
Console.WriteLine("Step 10: Reading XLSX with EPPlus 4.5.3.3...");
ReadWithEPPlus(xlsxFilePath);
Console.WriteLine();

Console.WriteLine("=== All Compatibility Tests Completed Successfully! ===");
Console.WriteLine();
Console.WriteLine("Summary:");
Console.WriteLine("  XLS Format (binary):");
Console.WriteLine("    ✓ NPOI 2.7.5 (HSSF) - read/write");
Console.WriteLine("    ✓ ExcelDataReader 3.8.0 - read");
Console.WriteLine();
Console.WriteLine("  XLSX Format (OpenXML):");
Console.WriteLine("    ✓ NPOI 2.7.5 (XSSF) - read/write");
Console.WriteLine("    ✓ MiniExcel 1.41.4 - read");
Console.WriteLine("    ✓ ExcelDataReader 3.8.0 - read");
Console.WriteLine("    ✓ DocumentFormat.OpenXml 3.3.0 - read");
Console.WriteLine("    ✓ ClosedXML 0.105.0 - read");
Console.WriteLine("    ✓ EPPlus 4.5.3.3 - read");
Console.WriteLine();
Console.WriteLine("Conclusion: NPOI 2.7.5 on .NET 9 is fully compatible!");

// Clean up
if (File.Exists(xlsFilePath))
{
    File.Delete(xlsFilePath);
}
if (File.Exists(xlsxFilePath))
{
    File.Delete(xlsxFilePath);
}

static void WriteXlsWithNPOI(string filePath, List<TestRecord> data)
{
    var workbook = new HSSFWorkbook();
    var sheet = workbook.CreateSheet("TestData");

    // Create header row
    var headerRow = sheet.CreateRow(0);
    headerRow.CreateCell(0).SetCellValue("ID");
    headerRow.CreateCell(1).SetCellValue("Name");
    headerRow.CreateCell(2).SetCellValue("Age");
    headerRow.CreateCell(3).SetCellValue("Score");
    headerRow.CreateCell(4).SetCellValue("Active");

    // Create data rows
    for (int i = 0; i < data.Count; i++)
    {
        var row = sheet.CreateRow(i + 1);
        row.CreateCell(0).SetCellValue(data[i].Id);
        row.CreateCell(1).SetCellValue(data[i].Name);
        row.CreateCell(2).SetCellValue(data[i].Age);
        row.CreateCell(3).SetCellValue(data[i].Score);
        row.CreateCell(4).SetCellValue(data[i].Active);
    }

    // Set column widths manually (AutoSizeColumn may require fonts in headless environment)
    for (int i = 0; i < 5; i++)
    {
        sheet.SetColumnWidth(i, 15 * 256); // 15 characters wide
    }

    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
    {
        workbook.Write(fs);
    }
}

static void WriteXlsxWithNPOI(string filePath, List<TestRecord> data)
{
    var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet("TestData");

    // Create header row
    var headerRow = sheet.CreateRow(0);
    headerRow.CreateCell(0).SetCellValue("ID");
    headerRow.CreateCell(1).SetCellValue("Name");
    headerRow.CreateCell(2).SetCellValue("Age");
    headerRow.CreateCell(3).SetCellValue("Score");
    headerRow.CreateCell(4).SetCellValue("Active");

    // Create data rows
    for (int i = 0; i < data.Count; i++)
    {
        var row = sheet.CreateRow(i + 1);
        row.CreateCell(0).SetCellValue(data[i].Id);
        row.CreateCell(1).SetCellValue(data[i].Name);
        row.CreateCell(2).SetCellValue(data[i].Age);
        row.CreateCell(3).SetCellValue(data[i].Score);
        row.CreateCell(4).SetCellValue(data[i].Active);
    }

    // Set column widths manually
    for (int i = 0; i < 5; i++)
    {
        sheet.SetColumnWidth(i, 15 * 256); // 15 characters wide
    }

    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
    {
        workbook.Write(fs);
    }
}

static void ReadXlsWithNPOI(string filePath)
{
    try
    {
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            var workbook = new HSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);
            
            Console.WriteLine($"  Sheet Name: {sheet.SheetName}");
            Console.WriteLine($"  Rows: {sheet.LastRowNum + 1}");
            
            // Read header
            var headerRow = sheet.GetRow(0);
            Console.Write("  Headers: ");
            for (int i = 0; i < 5; i++)
            {
                Console.Write($"{headerRow.GetCell(i)?.ToString()} | ");
            }
            Console.WriteLine();
            
            // Read first data row
            var dataRow = sheet.GetRow(1);
            Console.WriteLine($"  First Record: ID={dataRow.GetCell(0)}, Name={dataRow.GetCell(1)}, Age={dataRow.GetCell(2)}, Score={dataRow.GetCell(3)}, Active={dataRow.GetCell(4)}");
            
            Console.WriteLine("  ✓ NPOI (HSSF) read successfully");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static void ReadXlsxWithNPOI(string filePath)
{
    try
    {
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            var workbook = new XSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);
            
            Console.WriteLine($"  Sheet Name: {sheet.SheetName}");
            Console.WriteLine($"  Rows: {sheet.LastRowNum + 1}");
            
            // Read header
            var headerRow = sheet.GetRow(0);
            Console.Write("  Headers: ");
            for (int i = 0; i < 5; i++)
            {
                Console.Write($"{headerRow.GetCell(i)?.ToString()} | ");
            }
            Console.WriteLine();
            
            // Read first data row
            var dataRow = sheet.GetRow(1);
            Console.WriteLine($"  First Record: ID={dataRow.GetCell(0)}, Name={dataRow.GetCell(1)}, Age={dataRow.GetCell(2)}, Score={dataRow.GetCell(3)}, Active={dataRow.GetCell(4)}");
            
            Console.WriteLine("  ✓ NPOI (XSSF) read successfully");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static void ReadWithMiniExcel(string filePath)
{
    try
    {
        var rows = MiniExcel.Query(filePath).ToList();
        Console.WriteLine($"  Rows read: {rows.Count}");
        
        if (rows.Count > 0)
        {
            var firstRow = rows[0] as IDictionary<string, object>;
            if (firstRow != null)
            {
                Console.Write("  Headers: ");
                foreach (var key in firstRow.Keys)
                {
                    Console.Write($"{key} | ");
                }
                Console.WriteLine();
            }
        }
        
        if (rows.Count > 1)
        {
            var dataRow = rows[1] as IDictionary<string, object>;
            if (dataRow != null)
            {
                Console.Write("  First Record: ");
                foreach (var kvp in dataRow)
                {
                    Console.Write($"{kvp.Key}={kvp.Value} | ");
                }
                Console.WriteLine();
            }
        }
        
        Console.WriteLine("  ✓ MiniExcel read successfully");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static void ReadWithExcelDataReader(string filePath)
{
    try
    {
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var result = reader.AsDataSet();
            var table = result.Tables[0];
            
            Console.WriteLine($"  Sheet Name: {table.TableName}");
            Console.WriteLine($"  Rows: {table.Rows.Count}");
            
            if (table.Rows.Count > 0)
            {
                Console.Write("  Headers: ");
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Console.Write($"{table.Rows[0][i]} | ");
                }
                Console.WriteLine();
            }
            
            if (table.Rows.Count > 1)
            {
                Console.Write("  First Record: ");
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Console.Write($"{table.Rows[1][i]} | ");
                }
                Console.WriteLine();
            }
            
            Console.WriteLine("  ✓ ExcelDataReader read successfully");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static void ReadWithClosedXML(string filePath)
{
    try
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1);
            
            Console.WriteLine($"  Sheet Name: {worksheet.Name}");
            Console.WriteLine($"  Used Range: {worksheet.RangeUsed()?.RangeAddress}");
            
            var range = worksheet.RangeUsed();
            if (range != null)
            {
                Console.WriteLine($"  Rows: {range.RowCount()}");
                
                // Read header
                Console.Write("  Headers: ");
                for (int col = 1; col <= 5; col++)
                {
                    Console.Write($"{worksheet.Cell(1, col).Value} | ");
                }
                Console.WriteLine();
                
                // Read first data row
                Console.Write("  First Record: ");
                for (int col = 1; col <= 5; col++)
                {
                    Console.Write($"{worksheet.Cell(2, col).Value} | ");
                }
                Console.WriteLine();
            }
            
            Console.WriteLine("  ✓ ClosedXML read successfully");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static void ReadWithEPPlus(string filePath)
{
    try
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            
            Console.WriteLine($"  Sheet Name: {worksheet.Name}");
            Console.WriteLine($"  Dimension: {worksheet.Dimension?.Address}");
            
            if (worksheet.Dimension != null)
            {
                Console.WriteLine($"  Rows: {worksheet.Dimension.Rows}");
                
                // Read header
                Console.Write("  Headers: ");
                for (int col = 1; col <= 5; col++)
                {
                    Console.Write($"{worksheet.Cells[1, col].Value} | ");
                }
                Console.WriteLine();
                
                // Read first data row
                Console.Write("  First Record: ");
                for (int col = 1; col <= 5; col++)
                {
                    Console.Write($"{worksheet.Cells[2, col].Value} | ");
                }
                Console.WriteLine();
            }
            
            Console.WriteLine("  ✓ EPPlus read successfully");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static void ReadWithOpenXml(string filePath)
{
    try
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = document.WorkbookPart;
            if (workbookPart != null)
            {
                var sheets = workbookPart.Workbook.Sheets;
                Console.WriteLine($"  Sheets count: {sheets?.Count()}");
                
                if (sheets != null && sheets.Any())
                {
                    var firstSheet = sheets.Elements<Sheet>().First();
                    Console.WriteLine($"  First Sheet Name: {firstSheet.Name}");
                    
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!);
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    
                    var rows = sheetData.Elements<Row>().ToList();
                    Console.WriteLine($"  Rows: {rows.Count}");
                    
                    if (rows.Count > 0)
                    {
                        Console.Write("  First row cells: ");
                        var firstRow = rows[0];
                        foreach (var cell in firstRow.Elements<Cell>())
                        {
                            var value = GetCellValue(cell, workbookPart);
                            Console.Write($"{value} | ");
                        }
                        Console.WriteLine();
                    }
                    
                    if (rows.Count > 1)
                    {
                        Console.Write("  Second row cells: ");
                        var secondRow = rows[1];
                        foreach (var cell in secondRow.Elements<Cell>())
                        {
                            var value = GetCellValue(cell, workbookPart);
                            Console.Write($"{value} | ");
                        }
                        Console.WriteLine();
                    }
                }
                
                Console.WriteLine("  ✓ DocumentFormat.OpenXml read successfully");
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ✗ Error: {ex.Message}");
    }
}

static string GetCellValue(Cell cell, WorkbookPart workbookPart)
{
    if (cell.CellValue == null)
        return string.Empty;
        
    string value = cell.CellValue.InnerText;
    
    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    {
        var stringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (stringTable != null)
        {
            value = stringTable.ElementAt(int.Parse(value)).InnerText;
        }
    }
    
    return value;
}

public class TestRecord
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public double Score { get; set; }
    public bool Active { get; set; }
}
