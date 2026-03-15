using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;


namespace CRBExcelAddIn
{

    public class DataManager
    {
        public DataTable GetSqlData(string connectionString, string sql)
        {
            // "Integrated Security=True" active l'authentification Windows (SSO)
            string fullConnString = connectionString + ";Integrated Security=True;App=OrbitAddIn";

            using (SqlConnection conn = new SqlConnection(fullConnString))
            {
                DataTable dt = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);
                conn.Open();
                adapter.Fill(dt); // Remplit le tableau avec les résultats de la base
                return dt;
            }
        }
    }
}
