# table2csv
A simple command line tool to export tables from MS Access mdb/accdb databases to csv files.

For each exported table a csv-file with the corresponding name is created.

## Usage
```
  -f, --file         Required. Input database to be processed.
  -t, --tablename    Required. Table(s) to be exported, comma-separated.
  -v, --verbose      (Default: False) Prints all messages to standard output.
  --help             Display this help screen.
```
### Examples
* Export a single table: `table2csv -f Northwind.accdb -t Customers`. This will create `Customers.csv` in the current directory.
* Export multiple tables: `table2csv -f Northwind.mdb -t Customers,Employees`.

### Connectionstring
The application uses the following connection string: `"Provider=Microsoft.ACE.OLEDB.12.0;data source={0}"`

## Prerequisites
This tool requires a MS Access installation (or the separate Microsoft Access Database Engine).
