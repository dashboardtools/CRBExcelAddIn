using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace CRBExcelAddIn
{


    public static class QueryParser
    {
        // Extrait "Fund" et "date" de {{Fund:string}} ou {{ValuationDate:date}}
        public static List<string> DetectParameters(string sql)
        {
            var parameters = new List<string>();
            var matches = Regex.Matches(sql, @"\{\{(?<param>.*?)\}\}");
            foreach (Match m in matches)
            {
                parameters.Add(m.Groups["param"].Value);
            }
            return parameters;
        }

        // Remplace les {{param}} par les valeurs saisies par l'utilisateur
        public static string BuildFinalSql(string sql, Dictionary<string, string> userValues)
        {
            foreach (var entry in userValues)
            {
                // Note: En prod, utilisez des paramètres SQL pour éviter l'injection
                sql = sql.Replace("{{" + entry.Key + "}}", entry.Value);
            }
            return sql;
        }
    }
}
