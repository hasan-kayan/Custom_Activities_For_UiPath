using System;
using System.Activities;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Linktera.Excel.Basics.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using static System.Net.Mime.MediaTypeNames;
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

         
            // Determine Turkish Character Set 
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            string filePath = filepath; // Assign filepath to a string variable
            Console.WriteLine("File path: " + filePath); // Concatenate with other strings

            // Create an Excel application object with error expectations
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to create Excel application.");
                throw;
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
                throw;
            }

            Console.WriteLine("Excel file opened successfully.");

            // Clean up Excel objects
            excelApp.Visible = true;
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);


            // Outputs / Code does not have output service
            return (ctx) => {
            };
        }

        #endregion
    }
}

