using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Drawing;
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

        public void WriteToExcel(List<UnitTestResult> resultList)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(FileName)))
            {
                var sheet = package.Workbook.Worksheets.Add("TestResult");
                var i = 1;
                foreach (var result in resultList)
                {
                    sheet.Cells[i, 1].Value = result.TestName;
                    sheet.Cells[i, 1].AutoFitColumns();
                    sheet.Cells[i, 2].Value = result.Outcome;
                    sheet.Cells[i, 2].AutoFitColumns();
                    if (result.Outcome.Equals(TestOutcome.Failed.ToString(),StringComparison.OrdinalIgnoreCase))
                    {
                        var fill = sheet.Cells[i, 2].Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.Red);
                    }
                    else
                    {
                        var fill = sheet.Cells[i, 2].Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.ForestGreen);
                    }
                    sheet.Cells[i, 3].Value = result.Message;
                    sheet.Cells[i, 4].Value = result.StrackTrace;
                    i++;
                }
                
                package.Save();
            }
            
        }
    }
}
