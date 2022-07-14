﻿using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Linq;

namespace AreaChartExample
{
    internal class Program
    {
        static void Main()
        {
            CreateFileWithAreaChart();
        }

        private static void CreateFileWithAreaChart()
        {
            var workbook = new XSSFWorkbook();
            var worksheet = workbook.CreateSheet();
            worksheet.CreateDrawingPatriarch();

            var header = worksheet.CreateRow(0);

            header.CreateCell(1).SetCellValue("One");
            header.CreateCell(2).SetCellValue("Ten");
            header.CreateCell(3).SetCellValue("Hundred");
            header.CreateCell(4).SetCellValue("Thousand");

            var row1 = worksheet.CreateRow(1);

            row1.CreateCell(0).SetCellValue("Title2");
            row1.CreateCell(1).SetCellValue(2);
            row1.CreateCell(2).SetCellValue(20);
            row1.CreateCell(3).SetCellValue(200);
            row1.CreateCell(4).SetCellValue(2000);

            var row2 = worksheet.CreateRow(2);

            row2.CreateCell(0).SetCellValue("Title1");
            row2.CreateCell(1).SetCellValue(1);
            row2.CreateCell(2).SetCellValue(10);
            row2.CreateCell(3).SetCellValue(100);
            row2.CreateCell(4).SetCellValue(1000);

            var anchor = worksheet.DrawingPatriarch.CreateAnchor(0, 0, 0, 0, 0, 4, 10, 14);
            var chart = worksheet.DrawingPatriarch.CreateChart(anchor);
            var chartData = chart.ChartDataFactory.CreateAreaChartData<double, double>();

            var xSeries = new CellRangeAddress(header.RowNum, header.RowNum, 1, 4);
            var ySeries1 = new CellRangeAddress(row1.RowNum, row1.RowNum, 1, 4);
            var ySeries2 = new CellRangeAddress(row2.RowNum, row2.RowNum, 1, 4);

            var series1 = chartData.AddSeries(
                            DataSources.FromNumericCellRange(worksheet, xSeries),
                            DataSources.FromNumericCellRange(worksheet, ySeries1));

            var series2 = chartData.AddSeries(
                            DataSources.FromNumericCellRange(worksheet, xSeries),
                            DataSources.FromNumericCellRange(worksheet, ySeries2));

            var axis = chart.GetAxis();
            var areaBottomAxis = axis.FirstOrDefault(x => x.Position == AxisPosition.Bottom) ??
                                          chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            var areaLeftAxis = axis.FirstOrDefault(x => x.Position == AxisPosition.Left) ??
                                        chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

            ((IValueAxis)areaLeftAxis).SetCrossBetween(AxisCrossBetween.Between);

            chart.Plot(chartData, areaBottomAxis, areaLeftAxis);

            using (var stream = new FileStream("areaChart.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream, leaveOpen: false);
            }
        }
    }
}
