using System;
using System.Activities;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.ExcelActivities.CloseWorkbook.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace LinkteraRobotics.ExcelActivities.CloseWorkbook.Activities
{
    [LocalizedDisplayName(nameof(Resources.CloseWorkbook_DisplayName))]
    [LocalizedDescription(nameof(Resources.CloseWorkbook_Description))]
    public class CloseWorkbook : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CloseWorkbook_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.CloseWorkbook_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        #endregion


        #region Constructors

        public CloseWorkbook()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Path == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Path)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var path = Path.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            await Task.Run(() =>
            {
                try
                {
                    // Start a new Excel application
                    Application excelApp = new Application();

                    // Find the workbook by name or path and close it without saving changes
                    foreach (Workbook workbook in excelApp.Workbooks)
                    {
                        if (workbook.FullName == path)
                        {
                            workbook.Close(false);
                            Marshal.ReleaseComObject(workbook);
                            break;
                        }
                    }

                    // Quit the Excel application
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    // Handle any exceptions here
                    throw new Exception("Error closing Excel workbook: " + ex.Message);
                }
            }, cancellationToken);




            // Outputs
            return (ctx) => {
            };
        }

        #endregion
    }
}

