using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trx2Excel.Model
{
    public class UnitTestResult
    {
        public string TestName { get; set; }
        public string Outcome { get; set; }
        public string Message { get; set; }
        public string StrackTrace { get; set; }

    }
}
