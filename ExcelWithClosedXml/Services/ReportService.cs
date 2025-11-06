using ClosedXML.Excel;
using ExcelWithClosedXml.Models;

namespace ExcelWithClosedXml.Services
{
    public class ReportService
    {
        public void GenerateFullReport(List<Sale> sales, string filePath)
        {
            using var workbook = new XLWorkbook();

            CreateSummaryWorksheet(workbook, sales);
            CreateDetailedWorksheet(workbook, sales);
            CreateCategoryWorksheet(workbook, sales);
            CreateSellerWorksheet(workbook, sales);

            workbook.SaveAs(filePath);
            Console.WriteLine($"✅ Report generated successfully: {filePath}");
        }

        private void CreateSummaryWorksheet(XLWorkbook workbook, List<Sale> sales)
        {
            var ws = workbook.Worksheets.Add("📊 SUMMARY");

            // Main Title
            ws.Cell("A1").Value = "SALES REPORT";
            ws.Range("A1:D1").Merge();
            ws.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(18)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Fill.SetBackgroundColor(XLColor.DarkBlue)
                .Font.SetFontColor(XLColor.White);

            // report data
            ws.Cell("A2").Value = $"Created in: {DateTime.Now:dd/MM/yyyy HH:mm}";
            ws.Range("A2:D2").Merge();
            ws.Cell("A2").Style
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Font.SetItalic(true);

            // Indicadores (KPIs)
            int row = 4;
            var kpis = new[]
            {
                ("Total Sales", sales.Count.ToString("N0")),
                ("Total Revenue", sales.Sum(v => v.Total).ToString("C2")),
                ("Average Ticket", sales.Average(v => v.Total).ToString("C2")),
                ("Products Sold", sales.Sum(v => v.Quantity).ToString("N0"))
            };

            foreach (var (label, value) in kpis)
            {
                ws.Cell(row, 1).Value = label;
                ws.Cell(row, 1).Style.Font.SetBold(true);

                ws.Cell(row, 2).Value = value;
                ws.Cell(row, 2).Style
                    .Font.SetFontSize(14)
                    .Font.SetFontColor(XLColor.DarkGreen);

                row++;
            }

            // Top 5 Products
            row += 2;
            ws.Cell(row, 1).Value = "🏆 TOP 5 SELLING PRODUCTS";
            ws.Cell(row, 1).Style.Font.SetBold(true).Font.SetFontSize(14);
            row++;

            var top5 = sales
                .GroupBy(v => v.Product)
                .Select(g => new { Product = g.Key, Revenue = g.Sum(v => v.Total) })
                .OrderByDescending(x => x.Revenue)
                .Take(5)
                .ToList();

            // Headers
            ws.Cell(row, 1).Value = "Product";
            ws.Cell(row, 2).Value = "Revenue";
            ws.Range(row, 1, row, 2).Style
                .Font.SetBold(true)
                .Fill.SetBackgroundColor(XLColor.LightGray);
            row++;

            foreach (var item in top5)
            {
                ws.Cell(row, 1).Value = item.Product;
                ws.Cell(row, 2).Value = item.Revenue;
                ws.Cell(row, 2).Style.NumberFormat.Format = "R$ #,##0.00"; // Currency format
                row++;
            }

            ws.Columns(1, 2).AdjustToContents();
        }

        private void CreateDetailedWorksheet(XLWorkbook workbook, List<Sale> sales)
        {
            var ws = workbook.Worksheets.Add("📋 Detailed");

            // Title
            ws.Cell("A1").Value = "DETAILED SALES REPORT";
            ws.Range("A1:H1").Merge();
            ws.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(16)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Fill.SetBackgroundColor(XLColor.DarkBlue)
                .Font.SetFontColor(XLColor.White);

            // Headers
            var headers = new[] { "ID", "Date", "Product", "Category", "Qty", "Unit Price", "Total", "Seller" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(3, i + 1);
                cell.Value = headers[i];
                cell.Style
                    .Font.SetBold(true)
                    .Fill.SetBackgroundColor(XLColor.LightBlue)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            }

            // Data
            int row = 4;
            foreach (var sale in sales)
            {
                ws.Cell(row, 1).Value = sale.Id;
                ws.Cell(row, 2).Value = sale.SaleDate;
                ws.Cell(row, 3).Value = sale.Product;
                ws.Cell(row, 4).Value = sale.Category;
                ws.Cell(row, 5).Value = sale.Quantity;
                ws.Cell(row, 6).Value = sale.UnitValue;
                ws.Cell(row, 7).Value = sale.Total;
                ws.Cell(row, 8).Value = sale.Seller;

                // Format
                ws.Cell(row, 2).Style.DateFormat.Format = "dd/mm/yyyy";
                ws.Cell(row, 6).Style.NumberFormat.Format = "R$ #,##0.00";
                ws.Cell(row, 7).Style.NumberFormat.Format = "R$ #,##0.00";

                
                if (row % 2 == 0)
                {
                    ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                row++;
            }

            // Totals line
            ws.Cell(row, 6).Value = "GRAND TOTAL:";
            ws.Cell(row, 6).Style.Font.SetBold(true);
            ws.Cell(row, 7).FormulaA1 = $"=SUM(G4:G{row - 1})";
            ws.Cell(row, 7).Style
                .Font.SetBold(true).Fill.SetBackgroundColor(XLColor.Yellow)
                .NumberFormat.Format = "R$ #,##0.00";


            // Automatic filters
            ws.Range($"A3:H{row}").SetAutoFilter();

            ws.Columns().AdjustToContents();
        }

        private void CreateCategoryWorksheet(XLWorkbook workbook, List<Sale> sales)
        {
            var ws = workbook.Worksheets.Add("📦 By Category");

            ws.Cell("A1").Value = "SALES BY CATEGORY";
            ws.Range("A1:D1").Merge();
            ws.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(16)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Fill.SetBackgroundColor(XLColor.DarkGreen)
                .Font.SetFontColor(XLColor.White);

            var byCategory = sales
                .GroupBy(v => v.Category)
                .Select(g => new
                {
                    Category = g.Key,
                    Quantity = g.Sum(v => v.Quantity),
                    Revenue = g.Sum(v => v.Total),
                    SalesCount = g.Count()
                })
                .OrderByDescending(x => x.Revenue)
                .ToList();

            ws.Cell("A3").Value = "Category";
            ws.Cell("B3").Value = "No. of Sales";
            ws.Cell("C3").Value = "Qty. Products";
            ws.Cell("D3").Value = "Revenue";
            ws.Range("A3:D3").Style
                .Font.SetBold(true)
                .Fill.SetBackgroundColor(XLColor.LightGreen);

            int row = 4;
            foreach (var item in byCategory)
            {
                ws.Cell(row, 1).Value = item.Category;
                ws.Cell(row, 2).Value = item.SalesCount;
                ws.Cell(row, 3).Value = item.Quantity;
                ws.Cell(row, 4).Value = item.Revenue;
                ws.Cell(row, 4).Style.NumberFormat.Format = "R$ #,##0.00";
                row++;
            }

            ws.Cell(row, 1).Value = "TOTAL";
            ws.Cell(row, 1).Style.Font.SetBold(true);
            ws.Cell(row, 2).FormulaA1 = $"=SUM(B4:B{row - 1})";
            ws.Cell(row, 3).FormulaA1 = $"=SUM(C4:C{row - 1})";
            ws.Cell(row, 4).FormulaA1 = $"=SUM(D4:D{row - 1})";
            ws.Range(row, 1, row, 4).Style
                .Font.SetBold(true)
                .Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(row, 4).Style.NumberFormat.Format = "R$ #,##0.00";

            ws.Columns().AdjustToContents();
        }

        private void CreateSellerWorksheet(XLWorkbook workbook, List<Sale> sales)
        {
            var ws = workbook.Worksheets.Add("👤 By Seller");

            ws.Cell("A1").Value = "PERFORMANCE BY SELLER";
            ws.Range("A1:D1").Merge();
            ws.Cell("A1").Style
                .Font.SetBold(true)
                .Font.SetFontSize(16)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Fill.SetBackgroundColor(XLColor.DarkOrange)
                .Font.SetFontColor(XLColor.White);

            var bySeller = sales
                .GroupBy(v => v.Seller)
                .Select(g => new
                {
                    Seller = g.Key,
                    SalesCount = g.Count(),
                    Revenue = g.Sum(v => v.Total),
                    AverageTicket = g.Average(v => v.Total)
                })
                .OrderByDescending(x => x.Revenue)
                .ToList();

            ws.Cell("A3").Value = "Seller";
            ws.Cell("B3").Value = "No. of Sales";
            ws.Cell("C3").Value = "Revenue";
            ws.Cell("D3").Value = "Average Ticket";
            ws.Range("A3:D3").Style
                .Font.SetBold(true)
                .Fill.SetBackgroundColor(XLColor.LightCoral);

            int row = 4;
            foreach (var item in bySeller)
            {
                ws.Cell(row, 1).Value = item.Seller;
                ws.Cell(row, 2).Value = item.SalesCount;
                ws.Cell(row, 3).Value = item.Revenue;
                ws.Cell(row, 4).Value = item.AverageTicket;

                ws.Cell(row, 3).Style.NumberFormat.Format = "R$ #,##0.00";
                ws.Cell(row, 4).Style.NumberFormat.Format = "R$ #,##0.00";
                row++;
            }

            ws.Columns().AdjustToContents();
        }
    }
}