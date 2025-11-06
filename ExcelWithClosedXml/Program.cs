using ExcelWithClosedXml.Data;
using ExcelWithClosedXml.Services;

Console.WriteLine("╔════════════════════════════════════════════╗");
Console.WriteLine("║  📊 EXCEL SALES REPORT GENERATOR           ║");
Console.WriteLine("╚════════════════════════════════════════════╝");
Console.WriteLine();


Console.Write("📦 Generating sales data... ");
var sales = SalesData.GenerateFictitiousSales(100);
Console.WriteLine($"✅ {sales.Count} sales generated!");

// Generate report
Console.Write("📄 Creating Excel report... ");
var service = new ReportService();
var filePath = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
    $"Sales_Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
);

service.GenerateFullReport(sales, filePath);

Console.WriteLine();
Console.WriteLine("═══════════════════════════════════════════");
Console.WriteLine($"✅ Report saved to:");
Console.WriteLine($"  {filePath}");
Console.WriteLine("═══════════════════════════════════════════");
Console.WriteLine();
Console.WriteLine("Press any key to exit...");
Console.ReadKey();