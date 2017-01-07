using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Trx2Excel.Model;
using Trx2Excel.Setting;

namespace Trx2Excel.TrxReaderUtil
{
    public class TrxReader
    {
        private string FileName { get; set; }
        public TrxReader(string fileName)
        {
            FileName = fileName;
        }

        public List<UnitTestResult> GetTestResults()
        {
            var resultList = new List<UnitTestResult>();
            var doc = new XmlDocument();
            doc.Load(FileName);
            var xmlNodeList = doc.GetElementsByTagName(NodeName.UnitTestResult);
            if (xmlNodeList != null)
            {
                foreach (XmlNode node in xmlNodeList)
                {
                    resultList.Add(GetResult(node));
                }
            }
            return resultList;

        }

        public UnitTestResult GetResult(XmlNode node)
        {
            var result = new UnitTestResult();
            result.TestName = node.Attributes?[NodeName.TestName]?.InnerText;
            result.Outcome = node.Attributes?[NodeName.Outcome]?.InnerText;
            var outcome = (TestOutcome)Enum.Parse(typeof(TestOutcome), result.Outcome, true);
            switch (outcome)
            {
                case TestOutcome.Failed:
                    var output = node.ChildNodes[GetNodeIndex(node, NodeName.Output)];
                    var errorInfo = output.ChildNodes[GetNodeIndex(output, NodeName.ErrorInfo)];
                    result.Message = errorInfo.ChildNodes[GetNodeIndex(errorInfo, NodeName.Message)]?.InnerText;
                    result.StrackTrace = errorInfo.ChildNodes[GetNodeIndex(node, NodeName.StackTrace)]?.InnerText;
                    break;
                case TestOutcome.Passed:
                    break;
                case TestOutcome.Skipped:
                    break;
            }
            return result;
        }

        public int GetNodeIndex(XmlNode node, string nodeName)
        {
            for (var i = 0; i < node.ChildNodes.Count; i++)
            {
                if (node.ChildNodes[i].Name.Equals(nodeName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return 0;
        }
    }
}
