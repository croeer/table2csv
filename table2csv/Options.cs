using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using CommandLine.Text;

namespace table2csv
{
    class Options
    {
        [Option('f', "file", Required = true, HelpText = "Input database to be processed.")]
        public string InputFile { get; set; }

        [OptionList('t', "tablename", Required = false, Separator = ',', HelpText = "Table(s) to be exported, comma-separated.", MutuallyExclusiveSet = "dump")]
        public IList<string> tableName { get; set; }

        [Option('d', "dump", Required = false, HelpText = "Dump the whole database (i.e. all tables)", MutuallyExclusiveSet = "dump")]
        public bool DumpAll { get; set; }

        [Option('l', "list", Required = false, HelpText = "List all table names.", MutuallyExclusiveSet = "dump")]
        public bool ListTables { get; set; }

        [Option('o', "output", Required = false, HelpText = "Ouput directory.")]
        public string OutputDir { get; set; }

        [Option('v', "verbose", DefaultValue = false, HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo("table2csv", "1.1.0"),
                Copyright = new CopyrightInfo("croeer\n", DateTime.Now.Year),
                AdditionalNewLineAfterOption = false,
                AddDashesToOption = true
            };
            help.AddPreOptionsLine("A simple tool to export tables from an MS Access mdb-Database to csv format. " +
                                   "For each table a separate file will be created.");
            help.AddOptions(this);
            return help + "\n";
        }

    }
}
