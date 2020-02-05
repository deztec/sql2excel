using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//
using System.IO;
using NDesk.Options;
using System.Data;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace sql2excel
{
    class Program
    {
        static int Verbosity;

        static int Main(string[] Arguments)
        {
            int HandleArgumentsResult = HandleArguments(Arguments);

            if (HandleArgumentsResult != 0)
            {
                return HandleArgumentsResult;
            }

            try
            {
                using (SqlConnection DBConnection = new SqlConnection(Settings.DBConnectionString))
                {
                    using (DataTable myDataTable = GetDataTableForQuery(DBConnection, Settings.Worksheet, Settings.SQLCommand, Settings.SQLCommandTimeOut))
                    {
                        if (myDataTable.Rows.Count > 0)
                        {
                            PopulateSheet(myDataTable);
                        }
                    }
                }

                Debug("{0} Done.", DateTime.Now.ToString());
                //Console.ReadKey();
            }//try
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception encountered: {1}", DateTime.Now.ToString(), ex.ToString());

                //Console.ReadKey();
                return 1;
            }

            return 0;
        }

        static int HandleArguments(string[] Arguments)
        {
            bool ShowHelp = false;

            OptionSet RequiredOptions = new OptionSet() {
                { "c|connection=", "Connection string for SQL server\r\nExample:\r\n\"Server=localhost; Trusted_Connection = True;\"",
                    v => Settings.DBConnectionString = v },
                { "q|query=", "SQL query to execute\r\nExample:\r\n\"SELECT * FROM dbo.DataTable\"",
                    v => Settings.SQLCommand = v },
                { "w|workbook=", "Path to Excel worbook to populate",
                    v => Settings.ExcelFileFullPath = v },
                { "s|worksheet=", "Name of Worksheet in workbook to insert data into",
                    v => Settings.Worksheet = v },
            };

            OptionSet OptionalOptions = new OptionSet() {
                { "m|max=", "Maximum number of seconds for SQL query to complete.  Default if not specified is 900 seconds (15 minutes).",
                    v => Settings.SQLCommandTimeOut = Convert.ToInt32(v) },
                { "t|type=", "Output type to use.  Valid options D, M, or T.  \r\n[D]erived - cell type and formatting determined based on data type of fields in result set  \r\n\r\n[M]atchLastRow - match the type and formatting of the last row in the worksheet.  This is helpful when styling/fonts need to match existing records.  \r\n\r\n[T]extOnly - outputs all data as Text cells.  Fastest output type.",
                    v => {
                        switch(v.ToUpper())
                        {
                            case "M":
                            case "MATCHLASTROW":
                                Settings.OutputType = OutputType.MATCHLASTROW;
                                break;
                            case "T":
                            case "TEXTONLY":
                                Settings.OutputType = OutputType.TEXTONLY;
                                break;
                            case "D":
                            case "DERIVED":
                            default:
                                Settings.OutputType = OutputType.DERIVED;
                                break;
                        } } },
                { "d|dateformat=", "For DERIVED output type only: Date format for date cells.  Defaults to \"YYYY-MM-DD\" if not specified.",
                    v => Settings.DefaultDateFormat = v },
                { "b|bitsasnumbers", "For DERIVED output type only: Output BIT fields as numbers [1|0] instead of [TRUE|FALSE]",
                    v => { if (v != null) Settings.OutputBitAsNumber = true; } },
                { "r|removelastrow", "Remove last row in existing worksheet",
                    v => { if (v != null) Settings.RemoveLastRow = true; } },
                { "n|newworkbook", "Create new workbook even if it already exists.\r\nWarning: you will lose existing data.",
                    v => { if (v != null) Settings.OverwriteWorkbook = true; } },
                { "o|ovewrite", "Overwrite worksheet if it already exists in the workbook.  Default behavior is to create if doesn't exist or append if it does.\r\nWarning: you will lose existing data.",
                    v => { if (v != null) Settings.OverwriteWorkSheet = true; } },
                { "h|headers", "Enable header output",
                    v => { if (v != null) Settings.OutputHeaders = true; } },
                { "v|verbose", "increase debug message verbosity",
                  v => { if (v != null) ++Verbosity; } },
            };

            try
            {
                RequiredOptions.Parse(Arguments);
                OptionalOptions.Parse(Arguments);

                //required arguments
                if (
                    String.IsNullOrWhiteSpace(Settings.DBConnectionString) ||
                    String.IsNullOrWhiteSpace(Settings.SQLCommand) ||
                    String.IsNullOrWhiteSpace(Settings.ExcelFileFullPath) ||
                    String.IsNullOrWhiteSpace(Settings.Worksheet)
                   )

                {
                    Console.WriteLine("***** Error:");
                    Console.WriteLine("***** Missing one or more required arguments.\r\n");
                    ShowHelp = true;
                }

                if (Settings.OverwriteWorkbook == true && Settings.OutputType == OutputType.MATCHLASTROW)
                {
                    Console.WriteLine("***** Error:");
                    Console.WriteLine("***** Cannot use [o]verwrite workbook option and [M]atchLastRow output types together.\r\n");
                    ShowHelp = true;
                }
                
            }
            catch (OptionException e)
            {
                Console.Write("{0}: ", System.AppDomain.CurrentDomain.FriendlyName);
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `{0} --help' for more information.", System.AppDomain.CurrentDomain.FriendlyName);
                return 1;
            }

            if (ShowHelp)
            {
                Program.ShowHelp(RequiredOptions, OptionalOptions);
                return 1;
            }

            Debug("DBConnectionString: \t\t{0}", Settings.DBConnectionString);
            Debug("SQLCommand: \t\t{0}", Settings.SQLCommand);
            Debug("ExcelFileFullPath: \t{0}", Settings.ExcelFileFullPath);
            Debug("Worksheet: \t{0}", Settings.Worksheet);
            Debug("OutputType: \t{0}", Settings.OutputType);
            Debug("DefaultDateFormat: \t{0}", Settings.DefaultDateFormat);
            Debug("OutputBitAsNumber: \t{0}", Settings.OutputBitAsNumber);
            Debug("RemoveLastRow: \t{0}", Settings.RemoveLastRow);
            Debug("OverwriteWorkbook: \t\t{0}", Settings.OverwriteWorkbook);
            Debug("OverwriteWorkSheet: \t\t{0}", Settings.OverwriteWorkSheet);
            Debug("OutputHeaders: \t\t{0}", Settings.OutputHeaders);
            Debug("\r\n");

            return 0;
        }

        static void ShowHelp(OptionSet RequiredOptions, OptionSet OptionalOptions)
        {
            Console.WriteLine("Usage: {0} [OPTIONS]", System.AppDomain.CurrentDomain.FriendlyName);
            Console.WriteLine("Executes a SQL Command and populates an Excel worksheet with the result set.");
            Console.WriteLine();
            Console.WriteLine("Required parameters:\r\n");
            RequiredOptions.WriteOptionDescriptions(Console.Out);
            Console.WriteLine("\r\nOptional parameters:\r\n");
            OptionalOptions.WriteOptionDescriptions(Console.Out);
            Console.WriteLine();
            Console.WriteLine("Example:\r\n {0} -c \"Server=localhost;Trusted_Connection=True;\" -q \"SELECT * FROM dbo.DataTable\" -w \"C:\\Temp\\MyWorkbook.xlsx\" -s \"Sheet1\"", System.AppDomain.CurrentDomain.FriendlyName);
        }

        static void Debug(string format, params object[] args)
        {
            if (Verbosity > 0)
            {
                Console.Write("# ");
                Console.WriteLine(format, args);
            }
        }

        static DataTable GetDataTableForQuery(SqlConnection DBConnection, string WorkSheetName, string SQL, int SQLTimeout)
        {
            SqlDataAdapter MyDataAdapter = new SqlDataAdapter(SQL, DBConnection);
            DataTable MyDataTable = new DataTable();

            MyDataAdapter.SelectCommand.CommandTimeout = SQLTimeout; //seconds -- default 900 = 15 min

            DateTime DataSetFillStart = DateTime.Now;
            Debug("{0} Filling dataset using query: \r\n\t{1}", DateTime.Now.ToString(), SQL);
            MyDataAdapter.Fill(MyDataTable);
            DateTime DataSetFillEnd = DateTime.Now;

            double DataSetFillTime = ((TimeSpan)(DataSetFillEnd - DataSetFillStart)).TotalSeconds;

            Debug("{0} Filled dataset with {1} records in {2:0.00} seconds. {3:0} R/s", DateTime.Now.ToString(), MyDataTable.Rows.Count, DataSetFillTime, MyDataTable.Rows.Count / DataSetFillTime);

            return MyDataTable;
        }
 
        static void PopulateSheet(DataTable MyDataTable)
        {
            IWorkbook Workbook;

            //Creates a blank workbook if it doesn't exist or if we're overwriting the workbook.
            if (!File.Exists(Settings.ExcelFileFullPath) || Settings.OverwriteWorkbook)
            {
                Debug("{0} Creating new workbook", DateTime.Now.ToString());
                Workbook = new XSSFWorkbook();
            }
            else
            {
                Debug("{0} Opening existing workbook {1}", DateTime.Now.ToString(), Settings.ExcelFileFullPath);

                using (FileStream Stream = new FileStream(Settings.ExcelFileFullPath, FileMode.Open, FileAccess.Read))
                {
                    Workbook = new XSSFWorkbook(Stream);
                }
            }

            using (FileStream Stream = new FileStream(Settings.ExcelFileFullPath, FileMode.Create, FileAccess.Write))
            {
                ISheet Worksheet = Workbook.GetSheet(Settings.Worksheet);
                
                if (Worksheet == null)
                {
                    Debug("{0} Did not find an existing worksheet with the name '{1}'", DateTime.Now.ToString(), Settings.Worksheet);
                    Debug("{0} Creating worksheet '{1}'", DateTime.Now.ToString(), Settings.Worksheet);
                    Worksheet = Workbook.CreateSheet(Settings.Worksheet);
                }
                else
                {
                    Debug("{0} Existing worksheet {1} found. Rows: {2}", DateTime.Now.ToString(), Settings.Worksheet, Worksheet.PhysicalNumberOfRows);
                }

                int NumRows = 0;
                int NumCols = MyDataTable.Columns.Count;

                List<TemplateField> TemplateFields = null;

                int CurrentRowNumber = 0;

                switch (Settings.OutputType)
                {
                    case OutputType.MATCHLASTROW:
                        //if the last row in the worksheet contains a sample row, build the List of TemplateFields using it:
                        TemplateFields = GetListOfTemplateFieldsFromSampleRow(Worksheet.GetRow(Worksheet.LastRowNum));

                        if (NumCols != TemplateFields.Count)
                        {
                            throw new MissingFieldException("Number of columns in dataset does not match number of columns in the sample row.");
                        }

                        break;
                    case OutputType.TEXTONLY:
                        break;

                    case OutputType.DERIVED:
                    default:
                        TemplateFields = GetListOfTemplateFieldsFromData(MyDataTable.Columns, Workbook);
                        break;
                }

                if (!Settings.OverwriteWorkSheet)
                {
                    CurrentRowNumber = Worksheet.LastRowNum;

                    if (Settings.RemoveLastRow == false && Worksheet.PhysicalNumberOfRows > 0)
                    {
                        CurrentRowNumber++;
                    }
                }
                else
                {
                    //fresh sheet:

                    //save the sheet order
                    int Order = Workbook.GetSheetIndex(Worksheet);

                    //remove the sheet and replace with a fresh one at the same Order as before
                    Workbook.RemoveSheetAt(Order);
                    Workbook.CreateSheet(Worksheet.SheetName);
                    Workbook.SetSheetOrder(Worksheet.SheetName, Order);

                    Worksheet = Workbook.GetSheetAt(Order);
                }

                DateTime InsertsStart = DateTime.Now;

                if (Settings.OutputHeaders)
                {
                    IRow Row = Worksheet.CreateRow(CurrentRowNumber);
                    int CurrentColumnNumber = 0;

                    foreach (DataColumn col in MyDataTable.Columns)
                    {
                        ICell Cell = Row.CreateCell(CurrentColumnNumber);

                        if (Settings.OutputType != OutputType.TEXTONLY)
                        {
                            Cell.CellStyle = TemplateFields[CurrentColumnNumber].CellStyle;
                        }

                        Cell.SetCellValue(col.ColumnName);
                        CurrentColumnNumber++;
                    }

                    CurrentRowNumber++;
                    NumRows++;
                }

                foreach (DataRow DataRow in MyDataTable.Rows)
                {
                    IRow Row = Worksheet.CreateRow(CurrentRowNumber);

                    int CurrentColumnNumber = 0;

                    foreach (DataColumn DataColumn in MyDataTable.Columns)
                    {
                        string Field = DataRow[DataColumn].ToString();

                        ICell Cell = Row.CreateCell(CurrentColumnNumber);

                        if (Settings.OutputType != OutputType.TEXTONLY)
                        {
                            Cell.CellStyle = TemplateFields[CurrentColumnNumber].CellStyle;

                            if (!String.IsNullOrWhiteSpace(Field))
                            {
                                if (TemplateFields[CurrentColumnNumber].CellType == CellType.Numeric)
                                {
                                    if (TemplateFields[CurrentColumnNumber].IsDateFormmated)
                                    {
                                        //it's a date
                                        Cell.SetCellValue(Convert.ToDateTime(Field));
                                    }
                                    else if (Double.TryParse(Field, out double DoubleResult))
                                    {
                                        //it's a number
                                        Cell.SetCellValue(DoubleResult);
                                    }
                                    else if (Boolean.TryParse(Field, out bool BoolResult))
                                    {
                                        //it's a bit field
                                        Cell.SetCellValue(Convert.ToInt32(Convert.ToBoolean(Field)));
                                    }
                                    else
                                    {
                                        //numeric, but not a double, not a boolean, and not a date...treat it as text.
                                        Cell.SetCellValue(Field);
                                    }
                                }
                                else if (TemplateFields[CurrentColumnNumber].CellType == CellType.Boolean)
                                {
                                    //it's a bit field
                                    Cell.SetCellValue(Convert.ToBoolean(Field));
                                }
                                else
                                {
                                    //non-numeric field - treat as text
                                    Cell.SetCellValue(Field);
                                }
                            }
                        }
                        else
                        {
                            Cell.SetCellValue(Field);
                        }

                        CurrentColumnNumber++;
                    }

                    CurrentRowNumber++;
                    NumRows++;

                    if (NumRows % 1000 == 0)
                    {
                        Debug("{0} Output row count: {1}", DateTime.Now.ToString(), NumRows);
                    }
                }

                DateTime InsertsEnd = DateTime.Now;
                double ExcelFillTime = ((TimeSpan)(InsertsEnd - InsertsStart)).TotalSeconds;
                Debug("{0} Filled workbook with {1} records in {2:0.00} seconds. {3:0} R/s", DateTime.Now.ToString(), NumRows, ExcelFillTime, NumRows / ExcelFillTime);

                DateTime WriteStart = DateTime.Now;

                Debug("{0} Writing workbook", DateTime.Now.ToString());

                Workbook.Write(Stream);
                Workbook.Close();
                Stream.Close();

                Debug("{0} Closing workbook", DateTime.Now.ToString());

                DateTime WriteEnd = DateTime.Now;

                double ExcelWriteTime = ((TimeSpan)(WriteEnd - WriteStart)).TotalSeconds;

                if (ExcelWriteTime == 0)
                {
                    ExcelWriteTime = 1;
                }

                Debug("{0} Wrote workbook with {1} records {2:0.00} seconds. {3:0} R/s", DateTime.Now.ToString(), Worksheet.PhysicalNumberOfRows, ExcelWriteTime, NumRows / ExcelWriteTime);
            }
        }

        static List<TemplateField> GetListOfTemplateFieldsFromSampleRow(IRow TemplateRow)
        {
            List<TemplateField> TemplateFields = new List<TemplateField>();

            if (TemplateRow != null)
            {
                foreach (ICell Cell in TemplateRow)
                {
                    if (Cell.CellType == CellType.Numeric)
                    {
                        TemplateFields.Add(new TemplateField
                        {
                            CellStyle = Cell.CellStyle,
                            CellType = Cell.CellType,
                            IsDateFormmated = DateUtil.IsCellDateFormatted(Cell)
                        });
                    }
                    else
                    {
                        TemplateFields.Add(new TemplateField
                        {
                            CellStyle = Cell.CellStyle,
                            CellType = Cell.CellType,
                            IsDateFormmated = false
                        });
                    }
                }
            }

            return TemplateFields;
        }

        static List<TemplateField> GetListOfTemplateFieldsFromData(DataColumnCollection DataColumns, IWorkbook Workbook)
        {
            List<TemplateField> TemplateFields = new List<TemplateField>();

            foreach (DataColumn Column in DataColumns)
            {
                ICellStyle CellStyle = Workbook.CreateCellStyle();

                switch (Type.GetTypeCode(Column.DataType))
                {
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.UInt16:
                    case TypeCode.UInt32:
                    case TypeCode.UInt64:
                    case TypeCode.Double:
                    case TypeCode.Decimal:

                        TemplateFields.Add(new TemplateField
                        {
                            CellStyle = null,
                            CellType = CellType.Numeric,
                            IsDateFormmated = false
                        });

                        break;
                    case TypeCode.Boolean:

                        TemplateFields.Add(new TemplateField
                        {
                            CellStyle = null,
                            CellType = (Settings.OutputBitAsNumber) ? CellType.Numeric : CellType.Boolean,
                            IsDateFormmated = false
                        });

                        break;
                    case TypeCode.DateTime:

                        CellStyle.DataFormat = Workbook.CreateDataFormat().GetFormat(Settings.DefaultDateFormat);
                        
                        TemplateFields.Add(new TemplateField
                        {
                            CellStyle = CellStyle,
                            CellType = CellType.Numeric,
                            IsDateFormmated = true
                        });
                        break;
                    case TypeCode.Char:
                    case TypeCode.String:
                    default:
                        TemplateFields.Add(new TemplateField
                        {
                            CellStyle = null,
                            CellType = CellType.String,
                            IsDateFormmated = false
                        });

                        break;
                }
            }

            return TemplateFields;
        }
    }

    class Settings
    {
        public static string DBConnectionString;
        public static string SQLCommand;
        public static int SQLCommandTimeOut = 900;
        public static string ExcelFileFullPath;
        public static string Worksheet;

        public static bool OverwriteWorkbook = false;
        public static bool OverwriteWorkSheet = false;
        public static bool OutputHeaders = false;
        public static OutputType OutputType = OutputType.DERIVED;
        public static bool RemoveLastRow = false;

        public static string DefaultDateFormat = "yyyy-MM-dd";
        public static bool OutputBitAsNumber = false;
    }

    class TemplateField
    {
        public ICellStyle CellStyle;
        public CellType CellType;
        public bool IsDateFormmated;
    }

    enum OutputType
    {
        MATCHLASTROW,
        DERIVED,
        TEXTONLY
    }
}
