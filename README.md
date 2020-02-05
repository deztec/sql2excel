# sql2excel
.NET command line tool to populate Excel (xlsx) files from a SQL server query without Excel having to be installed.

Usage: sql2excel.exe [OPTIONS]
Executes a SQL Command and populates an Excel worksheet with the result set.

Required parameters:

  -c, --connection=VALUE     Connection string for SQL server
                               Example:
                               "Server=localhost; Trusted_Connection = True;"
  -q, --query=VALUE          SQL query to execute
                               Example:
                               "SELECT * FROM dbo.DataTable"
  -w, --workbook=VALUE       Path to Excel worbook to populate
  -s, --worksheet=VALUE      Name of Worksheet in workbook to insert data into

Optional parameters:

  -m, --max=VALUE            Maximum number of seconds for SQL query to
                               complete.  Default if not specified is 900
                               seconds (15 minutes).
  -t, --type=VALUE           Output type to use.  Valid options D, M, or T.

                               [D]erived - cell type and formatting determined
                               based on data type of fields in result set

                               [M]atchLastRow - match the type and formatting
                               of the last row in the worksheet.  This is
                               helpful when styling/fonts need to match
                               existing records.

                               [T]extOnly - outputs all data as Text cells.
                               Fastest output type.
  -d, --dateformat=VALUE     For DERIVED output type only: Date format for
                               date cells.  Defaults to "YYYY-MM-DD" if not
                               specified.
  -b, --bitsasnumbers        For DERIVED output type only: Output BIT fields
                               as numbers [1|0] instead of [TRUE|FALSE]
  -r, --removelastrow        Remove last row in existing worksheet
  -n, --newworkbook          Create new workbook even if it already exists.
                               Warning: you will lose existing data.
  -o, --ovewrite             Overwrite worksheet if it already exists in the
                               workbook.  Default behavior is to create if
                               doesn't exist or append if it does.
                               Warning: you will lose existing data.
  -h, --headers              Enable header output
  -v, --verbose              increase debug message verbosity

Example:

sql2excel.exe -c "Server=localhost;Trusted_Connection=True;" -q "SELECT * FROM dbo.DataTable" -w "C:\Temp\MyWorkbook.xlsx" -s "Sheet1"
