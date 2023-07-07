using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.Copy.Sheet.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;

namespace LinkteraRobotics.Copy.Sheet.Activities
{
    [LocalizedDisplayName(nameof(Resources.CopyExcelSheetValue_DisplayName))]
    [LocalizedDescription(nameof(Resources.CopyExcelSheetValue_Description))]
    public class CopyExcelSheetValue : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopyExcelSheetValue_Filepath_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopyExcelSheetValue_Filepath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Filepath { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopyExcelSheetValue_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopyExcelSheetValue_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        #endregion


        #region Constructors

        public CopyExcelSheetValue()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Filepath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Filepath)));
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var sheetname = Sheetname.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            ///

            string excelFilePath = filepath;
            string sheetName = sheetname;


            // Excel uygulamas�n� ba�lat
            Console.WriteLine("Starting Excel Application...");
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            // Aktif �al��ma kitab�n� ve �al��ma sayfas�n� al
            Console.WriteLine("Active Workbook Detection");
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = (Worksheet)workbook.Sheets[sheetName];

            try
            {
                Console.WriteLine("Selecting all cells...");
                // T�m h�creleri se� ve kopyala
                Excel.Range cells = worksheet.Cells;
                cells.Select();
                cells.Copy();

                // Yeni bir �al��ma sayfas� ekle
                Console.WriteLine("Creating new sheet");
                Excel.Worksheet newWorksheet = (Worksheet)workbook.Sheets.Add(After: workbook.ActiveSheet);

                // Yap��t�rma i�lemini ger�ekle�tir
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
                // Excel uygulamas�n� kapat ve kaynaklar� serbest b�rak
                workbook.Close();
                excelApp.Quit();
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

