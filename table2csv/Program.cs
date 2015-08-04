using CommandLine;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace table2csv
{
    class Program
    {
        static void Main(string[] args)
        {

            var options = new Options();

            /*
            var result = Parser.Default.ParseArguments(args, options);
            if (!result)
            {
                return;
            }
            */

            if (!new Parser(new ParserSettings
            {
                MutuallyExclusive = true,
                CaseSensitive = false,
                HelpWriter = Console.Error
            }).ParseArguments(args, options))
            {
                return;
            }

            // check mutually exclusive options
            if (!(options.DumpAll || options.tableName != null || options.ListTables))
            {
                Console.WriteLine(options.GetUsage());
                return;
            }

            if (options.Verbose)
                Console.WriteLine("");

            var fileName = options.InputFile;

            if (!File.Exists(fileName))
            {
                Console.WriteLine("File does not exist: {0}.", fileName);
                return;
            }

            if ((options.OutputDir != null) && ! Directory.Exists(options.OutputDir))
            {
                Console.WriteLine("Directory does jnot exist: {0}.", options.OutputDir);
                return;
            }

            string connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;data source={0}", fileName);

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

                if (options.ListTables)
                {
                    if (options.Verbose)
                        Console.WriteLine("Listing all tables from {0}", fileName);

                    var tableNameList = getAllTableNames(conn);
                    if (tableNameList != null)
                    {
                        foreach (var name in tableNameList)
                        {
                            Console.WriteLine(name);
                        }
                    }
                    return;
                }
                
                if (options.DumpAll)
                {
                    options.tableName = getAllTableNames(conn);
                }

                if (options.tableName == null)
                {
                    Console.WriteLine("Error determining table names. Aborting.");
                    return;
                }

                ExportTables(options, conn);
            }

        }

        private static void ExportTables(Options options, OleDbConnection conn)
        {
            foreach (var tableName in options.tableName)
            {
                Console.WriteLine("Exporting table {0}", tableName);

                var results = GetResults(conn, tableName);

                WriteDatatableToFile(options, results, tableName);
            }
        }

        private static void WriteDatatableToFile(Options options, DataTable results, string tableName)
        {
            StringBuilder sb = new StringBuilder();

            if (options.Verbose)
                Console.WriteLine("Parsing {0} rows, {1} columns", results.Rows.Count, results.Columns.Count);

            IEnumerable<string> columnNames = results.Columns.Cast<DataColumn>().
                Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in results.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }

            try
            {
                var outputPath = "{0}.csv";

                if (options.OutputDir != null)
                    outputPath = string.Format("{0}\\{1}", options.OutputDir, outputPath);

                var outputFileName = string.Format(outputPath, tableName);

                if (options.Verbose)
                    Console.WriteLine("Saving to {0}.", outputFileName);

                File.WriteAllText(outputFileName, sb.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing to file {0}.csv\n{1}", tableName, ex.Message);
            }
        }

        private static DataTable GetResults(OleDbConnection conn, string tableName)
        {
            DataTable results = new DataTable();

            try
            {
                OleDbCommand cmd = new OleDbCommand(String.Format("SELECT * FROM {0}", tableName), conn);
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

                adapter.Fill(results);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading table {0}: {1}", tableName, ex.Message);
                return results;
            }
            return results;
        }

        private static IList<string> getAllTableNames(OleDbConnection conn)
        {
            try
            {
                DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });

                return schemaTable.Rows.OfType<DataRow>().Select(r => r.ItemArray[2].ToString()).ToList();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading tables: {0}", ex.Message);
            }

            return null;
        }



    }
}
