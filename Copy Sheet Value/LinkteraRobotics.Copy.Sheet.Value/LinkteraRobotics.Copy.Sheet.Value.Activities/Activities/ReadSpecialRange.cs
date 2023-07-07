using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Copy.Sheet.Value.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;

using DataTable = System.Data.DataTable;
using System.Runtime.InteropServices;

namespace LinkteraRobotics.Copy.Sheet.Value.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadSpecialRange_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadSpecialRange_Description))]
    public class ReadSpecialRange : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadSpecialRange_Filepath_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadSpecialRange_Filepath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Filepath { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadSpecialRange_TargetSheet_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadSpecialRange_TargetSheet_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> TargetSheet { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadSpecialRange_TargetRange_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadSpecialRange_TargetRange_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> TargetRange { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadSpecialRange_Output_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadSpecialRange_Output_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> Output { get; set; }

        #endregion


        #region Constructors

        public ReadSpecialRange()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Filepath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Filepath)));
            if (TargetSheet == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(TargetSheet)));
            if (TargetRange == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(TargetRange)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var targetsheet = TargetSheet.Get(context);
            var targetrange = TargetRange.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            ///

            string excelFilePath = filepath;
            string sheetName = targetsheet;

            // Create an Excel application object
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Check if the workbook is already open
            Workbook workbook = null;
            foreach (Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.Equals(filepath, StringComparison.OrdinalIgnoreCase))
                {
                    workbook = wb;
                    break;
                }
            }

            // If the workbook is not open, open it

            if (workbook == null)
            {
                workbook = excelApp.Workbooks.Open(filepath);
            }

            // Get the worksheet
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets[sheetName] as Microsoft.Office.Interop.Excel.Worksheet;
            if (worksheet == null)
            {
                workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                throw new Exception($"Worksheet '{sheetName}' not found.");
            }

            try
            {
                // Tüm hücreleri seç ve kopyala
                Microsoft.Office.Interop.Excel.Range cells = worksheet.Cells;
                cells.Select();
                cells.Copy();

                // Yeni bir çalýþma sayfasý ekle
                Microsoft.Office.Interop.Excel.Worksheet newWorksheet = workbook.Sheets.Add(After: workbook.ActiveSheet) as Microsoft.Office.Interop.Excel.Worksheet;

                // Yapýþtýrma iþlemini gerçekleþtir
                Microsoft.Office.Interop.Excel.Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);

                Console.WriteLine("Kopyalama ve yapýþtýrma iþlemi tamamlandý.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata: " + ex.Message);
            }








            // Outputs
            return (ctx) => {
                Output.Set(ctx, null);
            };
        }

        #endregion
    }
}

