# table2csv
A simple command line tool to export tables from MS Access mdb/accdb databases to csv files.

For each exported table a csv-file with the corresponding name is created.

## Usage
```
  -f, --file         Required. Input database to be processed.
  -t, --tablename    Table(s) to be exported, comma-separated.
  -d, --dump         Dump the whole database (i.e. all tables)
  -l, --list         List all table names.
  -o, --output       Ouput directory.
  -v, --verbose      (Default: False) Prints all messages to standard output.
 --help             Display this help screen.  
```
The options `file`, `tablename` and `dump` are mutually exclusive, if one is present no other is available.
### Examples
* Export a single table: `table2csv -f Northwind.accdb -t Customers`. This will create `Customers.csv` in the current directory.
* Export multiple tables: `table2csv -f Northwind.mdb -t Customers,Employees`.
* Dump whole database to folder output: `table2csv -f Northwind.mdb -d -o output`.
* List all tables: `table2csv -f Northwind.mdb -d -l`.

### Connectionstring
The application uses the following connection string: `"Provider=Microsoft.ACE.OLEDB.12.0;data source={0}"`

## Prerequisites
This tool requires a MS Access installation (or the separate Microsoft Access Database Engine found [here](https://www.microsoft.com/de-de/download/details.aspx?id=13255)).

It also depends on the [Command Line Parser Library](https://github.com/gsscoder/commandline).
