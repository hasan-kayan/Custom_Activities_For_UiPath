using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotoics.Copy.Sheet.Value.Activities.Properties;
using Microsoft.Office.Interop.Excel;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.ComponentModel;

namespace LinkteraRobotoics.Copy.Sheet.Value.Activities
{
    [LocalizedDisplayName(nameof(Resources.CopySheetValue_DisplayName))]
    [LocalizedDescription(nameof(Resources.CopySheetValue_Description))]
    public class CopySheetValue : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheetValue_Filepath_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheetValue_Filepath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Filepath { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheetValue_Targetsheet_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheetValue_Targetsheet_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Targetsheet { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheetValue_Pastesheet_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheetValue_Pastesheet_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Pastesheet { get; set; }

        #endregion


        #region Constructors

        public CopySheetValue()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Filepath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Filepath)));
            if (Targetsheet == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Targetsheet)));
            if (Pastesheet == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Pastesheet)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var targetsheet = Targetsheet.Get(context);
            var pastesheet = Pastesheet.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            ///// Excel application object
            var excelApp = new Application();

            try
            {
                Workbook workbook = null;
                Worksheet sourceSheet = null;

                // Check if the Excel file is already open
                try
                {
                    workbook = excelApp.Workbooks.get_Item(filepath);
                    sourceSheet = (Worksheet)workbook.Sheets[targetsheet];
                }
                catch (Exception)
                {
                    // File is not open, open it
                    workbook = excelApp.Workbooks.Open(filepath);
                    sourceSheet = (Worksheet)workbook.Sheets[targetsheet];
                }

                // Add new target sheet
                Worksheet targetSheet = (Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                targetSheet.Name = targetsheet;

                // Copy source sheet data
                sourceSheet.Cells.Copy();
                // Paste values to target sheet
                targetSheet.Select();
                targetSheet.PasteSpecial();


                // Save and close workbook if it was opened during this execution
                if (sourceSheet == null)
                    workbook.Close(SaveChanges: true);
            }
            catch (Exception ex)
            {
                throw new Exception("Error processing Excel file. " + ex.Message);
            }
            try
            {
                // ...
            }
            catch (Exception ex)
            {
                throw new Exception("Error processing Excel file. " + ex.Message);
            }
            finally
            {
                // Close Excel application
                excelApp.Quit();
                workbook.Close(SaveChanges: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                workbook = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // Outputs
            return (ctx) =>
                {
                };
            }

            #endregion
        }
    }


