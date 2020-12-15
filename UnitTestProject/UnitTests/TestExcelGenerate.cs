using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Trx2Excel.ExcelUtils;
using Trx2Excel.TrxReaderUtil;

namespace UnitTestProject.UnitTests
{
    [TestClass]
    public class TestExcelGenerate
    {
        [TestMethod]
        public void TestExcelFileShouldGetCreated()
        {
            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var fileLocation = Path.Combine(assemblyFolder, "UnitTests", "MSTestSampleResult.trx");
            var excelLocation = Path.Combine(assemblyFolder, "UnitTests", "MSTestSampleResult.xlsx");
            var reader = new TrxReader(fileLocation);
            var resultList = reader.GetTestResults();
            var excelWriter = new ExcelWriter(excelLocation);
            excelWriter.WriteToExcel(resultList);
            Assert.IsTrue(File.Exists(excelLocation), "Fail to create Excel File");
        }
    }
}
