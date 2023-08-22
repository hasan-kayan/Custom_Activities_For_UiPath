using System;
using System.Activities;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.ExcelActivities.ReadCell.Activities.Properties;
using Microsoft.Office.Interop.Excel;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace LinkteraRobotics.ExcelActivities.ReadCell.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadCell_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadCell_Description))]
    public class ReadCell : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadCell_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadCell_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadCell_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadCell_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadCell_CellAddress_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadCell_CellAddress_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> CellAddress { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadCell_Value_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadCell_Value_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Value { get; set; }

        #endregion


        #region Constructors

        public ReadCell()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Path == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Path)));
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));
            if (CellAddress == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(CellAddress)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var path = Path.Get(context);
            var sheetname = Sheetname.Get(context);
            var celladdress = CellAddress.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////



            string cellValue = null;


            try
            {
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(path);

                // Find the desired sheet by name
                Worksheet worksheet = null;
                foreach (Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == sheetname)
                    {
                        worksheet = sheet;
                        break;
                    }
                }

                if (worksheet != null)
                {
                    // Read the cell value
                    Range cell = worksheet.Range[celladdress];
                    cellValue = cell.Value.ToString();

                    // Release COM objects
                    Marshal.ReleaseComObject(cell);
                    Marshal.ReleaseComObject(worksheet);
                }

                // Close the workbook and quit the application
                workbook.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                // Handle any exceptions here
                throw new Exception("Error reading cell value: " + ex.Message);
            }
















            // Outputs
            return (ctx) => {
                Value.Set(ctx, cellValue);
            };
        }

        #endregion
    }
}

