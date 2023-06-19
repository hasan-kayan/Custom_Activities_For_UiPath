using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Linktera.Excel.Basics.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Linktera.Excel.Basics.Activities
{
    [LocalizedDisplayName(nameof(Resources.OpenExcelFile_DisplayName))]
    [LocalizedDescription(nameof(Resources.OpenExcelFile_Description))]
    public class OpenExcelFile : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.OpenExcelFile_FilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.OpenExcelFile_FilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FilePath { get; set; }

        #endregion


        #region Constructors

        public OpenExcelFile()
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
            var filepath = FilePath.Get(context);

            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            string filePath = filepath; // Assign filepath to a string variable
            Console.WriteLine("File path: " + filePath); // Concatenate with other strings

            // Create an Excel application object
            Application excelApp = null;

            excelApp = new Application();

            // Open the workbook
            Workbook workbook = null;

            workbook = excelApp.Workbooks.Open(filePath);

            Console.WriteLine("Excel file opened successfully.");

            // Clean up Excel objects
            excelApp.Visible = true;
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);


            // Outputs
            return (ctx) => {
            };
        }

        #endregion
    }
}

