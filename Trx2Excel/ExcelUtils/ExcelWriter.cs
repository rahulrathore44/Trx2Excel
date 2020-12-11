using System;
using System.Collections.Generic;

using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using Trx2Excel.Model;
using Trx2Excel.Setting;

namespace Trx2Excel.ExcelUtils
{
    public class ExcelWriter
    {
        private string FileName { get; set; }
        public ExcelWriter(string fileName)
        {
            FileName = fileName;
        }

        /// <summary>
        /// Filter the data using the name space
        /// </summary>
        /// <param name="resultList"></param>
        /// <returns></returns>
        private Dictionary<string, List<UnitTestResult>> FilterTheDataWithNameSpace(List<UnitTestResult> resultList)
        {
            if (resultList?.Capacity < 1)
                throw new ArgumentException("No Test found in the trx report");

            HashSet<String> uniqueNameSpace = new HashSet<string>();
            Dictionary<string, List<UnitTestResult>> dataGroupByNameSpace = new Dictionary<string, List<UnitTestResult>>();

            foreach (UnitTestResult unitTestResult in resultList)
            {
                if (uniqueNameSpace.Add(unitTestResult.NameSpace))
                {
                    List<UnitTestResult> filteredList = resultList.FindAll((x) =>
                    {
                        return x.NameSpace.Equals(unitTestResult.NameSpace, StringComparison.OrdinalIgnoreCase);
                    });
                    dataGroupByNameSpace.Add(unitTestResult.NameSpace, filteredList);
                }
            }
            return dataGroupByNameSpace;
        }

        public void WriteToExcel(List<UnitTestResult> resultList)
        {
            var filteredData = FilterTheDataWithNameSpace(resultList);

            using (var package = new ExcelPackage(new System.IO.FileInfo(FileName)))
            {
                foreach (string nameSpace in filteredData.Keys)
                    AddSheetToExcel(filteredData, package, nameSpace);
                package.Save();
            }
            
        }

        private void AddSheetToExcel(Dictionary<string, List<UnitTestResult>> filteredData, ExcelPackage package, string nameSpace)
        {
            var sheet = package.Workbook.Worksheets.Add(nameSpace);
            sheet = CreateHeader(sheet);
            var i = 2;
            foreach (var result in filteredData[nameSpace])
            {
                sheet.Cells[i, 1].Value = result.TestName;
                sheet.Cells[i, 1].AutoFitColumns();
                sheet.Cells[i, 2].Value = result.Outcome;
                sheet.Cells[i, 2].AutoFitColumns();
                sheet.Cells[i, 3].Value = result.NameSpace;
                sheet.Cells[i, 3].AutoFitColumns();
                sheet.Cells[i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[i, 2].Style.Fill.BackgroundColor.SetColor(
                    result.Outcome.Equals(TestOutcome.Failed.ToString(), StringComparison.OrdinalIgnoreCase) ?
                    Color.Red :
                    Color.ForestGreen);
                sheet.Cells[i, 4].Value = result.Message;
                sheet.Cells[i, 5].Value = result.StackTrace;
                i++;
            }
        }

        public void AddChart(int pass, int fail, int skip)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(FileName)))
            {
                var sheet = package.Workbook.Worksheets.Add("Result Chart");
                var chart = sheet.Drawings.AddChart("Result Chart", eChartType.Pie);
                var barChart = sheet.Drawings.AddChart("Result Bar Chart", eChartType.BarStacked);
                
                AddDataToSheet(pass, fail, skip, sheet);

                chart.Title.Text = "Test Result Pie Chart";
                chart.SetPosition(1, 0, 3, 0);
                var ser = chart.Series.Add("B2:B4", "A2:A4");
                ser.Header = "Count";

                barChart.Title.Text = "Test Result Bar Chart";
                barChart.SetPosition(14, 0, 3, 0);
                var barSer = barChart.Series.Add("B2:B4", "A2:A4");
                barSer.Header = "Count";
                package.Save();
            }
        }

        private static void AddDataToSheet(int pass, int fail, int skip, ExcelWorksheet sheet)
        {
            sheet.Cells["A2"].Value = "Passed";
            sheet.Cells["A2"].AutoFitColumns();
            sheet.Cells["A3"].Value = "Failed";
            sheet.Cells["A3"].AutoFitColumns();
            sheet.Cells["A4"].Value = "Skipped";
            sheet.Cells["A4"].AutoFitColumns();
            sheet.Cells["B2"].Value = pass;
            sheet.Cells["B3"].Value = fail;
            sheet.Cells["B4"].Value = skip;
            sheet.Cells["A2,A3,A4,B2,B3,B4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A2,B2"].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
            sheet.Cells["A3,B3"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            sheet.Cells["A4,B4"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        }

        public ExcelWorksheet CreateHeader(ExcelWorksheet sheet)
        {
            string[] header = {"Test Name", "Status", "Name Space","Exception Message", "Stack Trace"};
            for (var i = 0; i < header.Length; i++)
            {
                sheet.Cells[1, i + 1].Value = header[i];
                sheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[1, i + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                sheet.Cells[1, i + 1].Style.Font.Bold = true;
                sheet.Cells[1, i + 1].AutoFitColumns();

            }
            return sheet;
        }
    }
}
