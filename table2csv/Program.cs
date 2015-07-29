using CommandLine;
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

            var options = new Options();

            var result = Parser.Default.ParseArguments(args, options);

            if (!result)
            {
                options.GetUsage();
                return;
            }

            if (options.Verbose)
                Console.WriteLine("");

            var fileName = options.InputFile;

            string connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;data source={0}", fileName);

            DataTable results = new DataTable();
            if (options.Verbose)
                Console.WriteLine("Using connectionstring {0}", connString);

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not open database {0}\n{1}.", fileName, ex.Message);
                    return;
                }

                foreach (var tableName in options.tableName)
                {
                    Console.WriteLine("Exporting table {0}", tableName);

                    try
                    {
                        OleDbCommand cmd = new OleDbCommand(String.Format("SELECT * FROM {0}", tableName), conn);
                        OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

                        adapter.Fill(results);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error reading table {0}: {1}", tableName, ex.Message);
                        continue;
                    }

                    if (options.Verbose)
                        Console.WriteLine("Parsing {0} rows, {1} columns", results.Rows.Count, results.Columns.Count);

                    StringBuilder sb = new StringBuilder();

                    IEnumerable<string> columnNames = results.Columns.Cast<DataColumn>().
                                                      Select(column => column.ColumnName);
                    sb.AppendLine(string.Join(",", columnNames));

                    foreach (DataRow row in results.Rows)
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(",", fields));
                    }

                    if (options.Verbose)
                        Console.WriteLine("Saving to {0}.csv", tableName);

                    try
                    {

                        File.WriteAllText(string.Format("{0}.csv", tableName), sb.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error writing to file {0}.csv\n{1}", tableName, ex.Message);
                    }
                }
            }

        }



    }
}
