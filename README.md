# table2csv
A simple command line tool to export tables from MS Access mdb/accdb databases to csv files.

## Usage
```
  -f, --file         Required. Input database to be processed.
  -t, --tablename    Required. Table(s) to be exported, comma-separated.
  -v, --verbose      (Default: False) Prints all messages to standard output.
  --help             Display this help screen.
```
### Examples
* Export a single table: `table2csv -f Nothwind.accdb -t Customers`
* Export multiple tables: `table2csv -f Nothwind.mdb -t Customers,Employees`

### Connectionstring
The application uses the following connection string: `"Provider=Microsoft.ACE.OLEDB.12.0;data source={0}"`

## Prerequisites
This tool requires a MS Access installation (or the separate Microsoft Access Database Engine).
