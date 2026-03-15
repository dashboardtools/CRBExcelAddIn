using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;


namespace CRBExcelAddIn
{
   

    public static class ExcelInjector
    {
        public static void Inject(DataTable data, string tableName)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.ListObject table = FindTable(app, tableName);

            if (table != null)
            {
                // 1. Désactiver les calculs pour aller 10x plus vite
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                // 2. Vider le contenu actuel sans supprimer le tableau lui-même
                if (table.DataBodyRange != null) table.DataBodyRange.Delete();

                // 3. Préparer le transfert par bloc (très performant)
                object[,] values = new object[data.Rows.Count, data.Columns.Count];
                for (int i = 0; i < data.Rows.Count; i++)
                    for (int j = 0; j < data.Columns.Count; j++)
                        values[i, j] = data.Rows[i][j];

                // 4. Coller les données d'un coup
                Excel.Range destination = table.HeaderRowRange.Offset[1, 0].Resize[data.Rows.Count, data.Columns.Count];
                destination.Value2 = values;

                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
        }

        private static Excel.ListObject FindTable(Excel.Application app, string name)
        {
            foreach (Excel.Worksheet ws in app.ActiveWorkbook.Worksheets)
                foreach (Excel.ListObject lo in ws.ListObjects)
                    if (lo.Name == name) return lo;
            return null;
        }
    }
}
