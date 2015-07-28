using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace table2csv
{
    class Program
    {
        static void Main(string[] args)
        {

            var fileName = "AutoTest500.mdb";
            var tableName = "DAM_ADRESSEN";

            string connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;data source={0}", fileName);

            /*
            string connString = 
                "Provider=Microsoft.ACE.OLEDB.12.0;persist security info=True;cache authentication=True;password=zebdepota;user id=zeb;data source=C:\\temp\\AutoTest500.mdb;Jet OLEDB:System database=C:\\dev\\dotnet\\trd-trunk\\bin\\x86_Debug\\trd.mda;Jet OLEDB:Engine Type=5";
            */

            DataTable results = new DataTable();

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                OleDbCommand cmd = new OleDbCommand(String.Format("SELECT * FROM {0}",tableName), conn);

                conn.Open();

                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(results);

            }

            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = results.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in results.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(string.Format("{0}.csv", tableName), sb.ToString());

        }



    }
}
