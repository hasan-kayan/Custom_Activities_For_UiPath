using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.LegacyV3.OpenExcel.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace LinkteraRobotics.LegacyV3.OpenExcel.Activities
{
    [LocalizedDisplayName(nameof(Resources.OpenExcel_DisplayName))]
    [LocalizedDescription(nameof(Resources.OpenExcel_Description))]
    public class OpenExcel : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.OpenExcel_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.OpenExcel_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        [LocalizedDisplayName(nameof(Resources.OpenExcel_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.OpenExcel_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        #endregion


        #region Constructors

        public OpenExcel()
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
            Console.WriteLine("Linktera Robotics");

            // Excel Application starts
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            Console.WriteLine("Workbook detection...");
            Excel.Workbook workbook = excelApp.Workbooks.Open(path);

            // Find the Worksheet with the specified name
            Excel.Worksheet worksheet = null;
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == sheetname)
                {
                    worksheet = ws;
                    break;
                }
            }

            // Check if the specified sheet exists, then activate it
            if (worksheet != null)
            {
                worksheet.Activate();
                Console.WriteLine("Sheet activated: " + worksheet.Name);
            }
            else
            {
                Console.WriteLine("Sheet not found: " + sheetname);
            }

            // Outputs
            return (ctx) => {
            };
        }

        #endregion
    }
}

