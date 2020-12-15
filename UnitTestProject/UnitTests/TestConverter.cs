using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Trx2Excel.ExcelUtils;
using Trx2Excel.Setting;
using Trx2Excel.TrxReaderUtil;

namespace Trx2Excel.UnitTests
{
    [TestClass]
    public class TestConverter
    {
        private readonly string ExpectedNameSpace = "Pickles.SpecFlow.Lab.Specs.BankAccountSpecs.TwoMoreScenariosTransferingFundsBetweenAccounts_OneFailngAndOneSuccedingFeature";

        [TestMethod]
        public void TestResultCount()
        {
            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var fileLocation = Path.Combine(assemblyFolder, "UnitTests", "MSTestSampleResult.trx");
            var reader = new TrxReader(fileLocation);
            var resultList = reader.GetTestResults();

            Assert.IsTrue(resultList.Count == 4, "Total Result should be 4");
            Assert.IsTrue(resultList.FindAll(test => test.Outcome == TestOutcome.Passed.ToString())?.Count == 3, "Total Result should be 3");

            Assert.IsTrue(resultList.FindAll(test => test.Outcome == TestOutcome.Failed.ToString())?.Count == 1, "Failed Result should be 1");
        }

        [TestMethod]
        public void TestNameSpace()
        {
            string TestName = "Transfer70WithCoverageFromOneAccountToAnother";
            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var fileLocation = Path.Combine(assemblyFolder, "UnitTests", "MSTestSampleResultTwo.trx");
            var reader = new TrxReader(fileLocation);
            var resultList = reader.GetTestResults();

            Assert.AreEqual(ExpectedNameSpace, resultList.First(test => test.TestName.Equals(TestName, StringComparison.OrdinalIgnoreCase)).NameSpace);
        }

        [TestMethod]
        public void TestExceptionMessage()
        {
            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var fileLocation = Path.Combine(assemblyFolder, "UnitTests", "MSTestSampleResultTwo.trx");
            var reader = new TrxReader(fileLocation);
            var resultList = reader.GetTestResults();

            Assert.IsNotNull(resultList.First(test => test.Outcome.Equals(TestOutcome.Failed.ToString(), StringComparison.OrdinalIgnoreCase)).StackTrace, "Stack trace should not be null for failed Test");
        }
    }
}
