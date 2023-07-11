using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Big.Data.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Excel = Microsoft.Office.Interop.Excel;

DataTable table = new DataTable();


namespace LinkteraRobotics.Read.Big.Data.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadLargeExcelData_Description))]
    public class ReadLargeExcelData : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_ExcelFilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_ExcelFilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> ExcelFilePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_OutputTable_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_OutputTable_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> OutputTable { get; set; }

        #endregion


        #region Constructors

        public ReadLargeExcelData()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));
            if (ExcelFilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ExcelFilePath)));
            if (Range == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Range)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var sheetname = Sheetname.Get(context);
            var excelfilepath = ExcelFilePath.Get(context);
            var range = Range.Get(context); // Has the same name


            // Added Variables Just because to not change structure

            string excelFilePath = excelfilepath;
            string sheetName = sheetname;

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            Console.WriteLine("READING BIG DATA PROCESS !");
            Console.WriteLine("Checking inputs: ");
            Console.WriteLine("Target Excel File Path: " + excelfilepath);
            Console.WriteLine("Target Excel Sheet: " + sheetName);
            Console.WriteLine("Target Range" +  range);


            // STARTING EXCEL 

            Console.WriteLine("Excel Application is Starting...");

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Console.WriteLine("Workbook detection...");
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = null;
            Excel.Worksheet newWorksheet = null;

            try
            {
                Console.WriteLine("Process is starting.");
                worksheet = (Excel.Worksheet)workbook.Sheets[sheetName];

                Console.WriteLine("Copying Cells");
                Excel.Range cells = worksheet.Cells;
                cells.Copy();
                Console.WriteLine("Cells Copied");

                Console.WriteLine("Creating a temporary worksheet.");
                newWorksheet = (Excel.Worksheet)workbook.Sheets.Add(After: workbook.ActiveSheet);
                newWorksheet.Name = "Copied";
                Console.WriteLine("Temporary sheet created ! ");

                Excel.Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);
                Console.WriteLine("Special pasted values to temproray worksheet.");
                Console.WriteLine("Checking range status.");

                if (newWorksheet != null)
                {
                    Console.WriteLine("Data Table Creating, data reading...");
                    Excel.Range excelRange;
                    if (string.IsNullOrWhiteSpace(range))
                    {
                        Console.WriteLine("All data reading");
                        excelRange = newWorksheet.UsedRange;
                    }
                    else
                    {
                        Console.WriteLine("The range : " + range + "is reading");
                        excelRange = newWorksheet.Range[range];
                    }

                    int rowCount = excelRange.Rows.Count;
                    int columnCount = excelRange.Columns.Count;

                    DataTable dataTable = new DataTable();

                    // Read column names
                    for (int col = 1; col <= columnCount; col++)
                    {
                        string columnName = excelRange.Cells[1, col]?.ToString() ?? $"Column{col}";

                        dataTable.Columns.Add(columnName);
                    }
                    Console.WriteLine("Column names read");

                    // Read data in smaller chunks (rows)
                    const int batchSize = 1000;
                    int remainingRows = rowCount - 1;
                    int startRow = 2;

                    while (remainingRows > 0)
                    {
                        int rowsToRead = Math.Min(batchSize, remainingRows);
                        int endRow = startRow + rowsToRead - 1;

                        Excel.Range dataRange = excelRange.Range[excelRange.Cells[startRow, 1], excelRange.Cells[endRow, columnCount]];
                        object[,] excelData = (object[,])dataRange.Value;
                        Console.WriteLine($"Read rows {startRow} to {endRow}");

                        // Transfer data to the DataTable
                        for (int row = 1; row <= rowsToRead; row++)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int col = 1; col <= columnCount; col++)
                            {
                                dataRow[col - 1] = excelData[row, col];
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                        Console.WriteLine($"Transferred rows {startRow} to {endRow}");

                        startRow += batchSize;
                        remainingRows -= rowsToRead;

                        OutputTable.Set(context, dataTable);
                    }

                  
                }
                else
                {
                    Console.WriteLine("New worksheet is null. Cannot proceed with data reading.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }
            finally
            {
                newWorksheet?.Activate();
                workbook.Close(false); // False states that dont save changes !!!!!!!!!
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }





            


            // Outputs
            return (ctx) => {
                OutputTable.Set(ctx, OutputTable);
            };
        }

        #endregion
    }
}

