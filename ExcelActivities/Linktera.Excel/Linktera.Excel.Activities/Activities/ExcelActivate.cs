using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Linktera.Excel.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Linktera.Excel.Activities
{
    [LocalizedDisplayName(nameof(Resources.ExcelActivate_DisplayName))]
    [LocalizedDescription(nameof(Resources.ExcelActivate_Description))]
    public class ExcelActivate : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ExcelActivate_FilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.ExcelActivate_FilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FilePath { get; set; }

        #endregion


        #region Constructors

        public ExcelActivate()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (FilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FilePath)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filePath = FilePath.Get(context);

            Application excelApp = null;
            try
            {
                excelApp = new Application();
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to create Excel application.");
                return null;
            }

            // Open the workbook
            Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to open the Excel file.");
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                return null;
            }

            Console.WriteLine("Excel file opened successfully.");

            // Outputs
            return (ctx) =>
            {
            };
        }


        #endregion
    }
}

