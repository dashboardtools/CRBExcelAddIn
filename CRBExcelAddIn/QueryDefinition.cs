using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRBExcelAddIn
{
    public class QueryDefinition
    {
        public string QueryId { get; set; }
        public string SqlQuery { get; set; }
        public string TargetSheet { get; set; }
        public string TargetTable { get; set; }
        public string ConnectionString { get; set; }
    }
}
