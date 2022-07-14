using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO;

namespace PieChartExample
{
    internal class Program
    {
        static void Main()
        {
            CreateFileWithPieChart();
        }
        
        private static void CreateFileWithPieChart()
        {
            var workbook = new XSSFWorkbook();
            var worksheet = workbook.CreateSheet();
            worksheet.CreateDrawingPatriarch();

            var header = worksheet.CreateRow(0);

            header.CreateCell(1).SetCellValue("Twenty");
            header.CreateCell(2).SetCellValue("Sixty");
            header.CreateCell(3).SetCellValue("Hundred");
            header.CreateCell(4).SetCellValue("Two hundreds");

            var row = worksheet.CreateRow(1);

            row.CreateCell(0).SetCellValue("Title2");
            row.CreateCell(1).SetCellValue(20);
            row.CreateCell(2).SetCellValue(60);
            row.CreateCell(3).SetCellValue(100);
            row.CreateCell(4).SetCellValue(200);

            var anchor = worksheet.DrawingPatriarch.CreateAnchor(0, 0, 0, 0, 0, 4, 6, 14);
            var chart = worksheet.DrawingPatriarch.CreateChart(anchor);
            var chartData = chart.ChartDataFactory.CreatePieChartData<string, double>();

            var xSeries = new CellRangeAddress(header.RowNum, header.RowNum, 1, 4);
            var ySeries = new CellRangeAddress(row.RowNum, row.RowNum, 1, 4);
            
            _ = chartData.AddSeries(
                            DataSources.FromStringCellRange(worksheet, xSeries),
                            DataSources.FromNumericCellRange(worksheet, ySeries));

            chart.Plot(chartData);

            using (var stream = new FileStream("pieChart.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream, leaveOpen: false);
            }
        }
    }
}
