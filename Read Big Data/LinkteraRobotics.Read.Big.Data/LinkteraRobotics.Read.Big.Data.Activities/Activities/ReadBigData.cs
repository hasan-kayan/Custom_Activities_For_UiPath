using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Big.Data.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace LinkteraRobotics.Read.Big.Data.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadBigData_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadBigData_Description))]
    public class ReadBigData : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_Paht_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_Paht_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Paht { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_SheetName_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_SheetName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> SheetName { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_OutTable_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_OutTable_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> OutTable { get; set; }

        #endregion


        #region Constructors

        public ReadBigData()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var path = Paht.Get(context);
            var sheetname = SheetName.Get(context);
            var range = Range.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            ///


            string excelFilePath = path;
            string sheetName = sheetname;


            Console.WriteLine("New Application Starting...");

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Console.WriteLine("Workbook detection...");
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = null;
            Excel.Worksheet newWorksheet = null;

            DataTable dataTable = null;


            try
            {
                worksheet = (Excel.Worksheet)workbook.Sheets[sheetName];

                Console.WriteLine("Copying Cells");
                Excel.Range cells = worksheet.Cells;
                cells.Copy();
                Console.WriteLine("Cells Copied");

                Console.WriteLine("Creating New Sheet");
                newWorksheet = (Excel.Worksheet)workbook.Sheets.Add(After: workbook.ActiveSheet);
                newWorksheet.Name = "Copied";
                Console.WriteLine("New Sheet named 'Copied' Created");

                Excel.Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);
                Console.WriteLine("Paste Special Completed");

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

                    

                    // Read column names
                    for (int col = 1; col <= columnCount; col++)
                    {
                        string columnName = excelRange.Cells[1, col]?.Value?.ToString() ?? $"Column{col}";
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
                OutTable.Set(ctx, dataTable);
            };
        }

        #endregion
    }
}







// Data table kodda try içinden kalktý try dýþýna tanýmlandý kullanýmýndan emin ol ve value methodunu araþtýr hatayý gider 
// sakýn sik sik yeni bi þey deneme vakit yok consolu aynen buraya eklememiz þart