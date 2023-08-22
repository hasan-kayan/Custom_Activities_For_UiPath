using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.ExcelActivities.CopySheet.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;

namespace LinkteraRobotics.ExcelActivities.CopySheet.Activities
{
    [LocalizedDisplayName(nameof(Resources.CopySheet_DisplayName))]
    [LocalizedDescription(nameof(Resources.CopySheet_Description))]
    public class CopySheet : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheet_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheet_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheet_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheet_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        #endregion


        #region Constructors

        public CopySheet()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Path == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Path)));
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var path = Path.Get(context);
            var sheetname = Sheetname.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            ///

            Console.WriteLine("Linktera Robotics");

            string excelFilePath = path;
            string sheetName = sheetname;


            // Start Excel Application
            Console.WriteLine("Starting Excel Application...");
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            // Detect Active Workbook & Worksheet
            Console.WriteLine("Active Workbook Detection");
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = (Worksheet)workbook.Sheets[sheetName];

            try
            {
                Console.WriteLine("Selecting all cells...");
                // Slect & Copy all cells
                Excel.Range cells = worksheet.Cells;
                cells.Select();
                cells.Copy();

                // Add new worksheet
                Console.WriteLine("Creating new sheet");
                Excel.Worksheet newWorksheet = (Worksheet)workbook.Sheets.Add(After: workbook.ActiveSheet);

                // PasteSpecial Process
                Console.WriteLine("Pasting data into new sheet...");
                Excel.Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);

                Console.WriteLine("Process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
            finally
            {
                // Relase Excel Objects

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }





            // Outputs
            return (ctx) => {
            };
        }

        #endregion
    }
}

