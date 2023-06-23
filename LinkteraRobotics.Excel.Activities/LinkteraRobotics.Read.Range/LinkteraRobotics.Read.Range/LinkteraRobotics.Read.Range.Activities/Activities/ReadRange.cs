using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Range.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text;

namespace LinkteraRobotics.Read.Range.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadRange_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadRange_Description))]
    public class ReadRange : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_ExcelRange_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_ExcelRange_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> ExcelRange { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_DataOut_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_DataOut_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<System.Data.DataTable> DataOut { get; set; }

        #endregion


        #region Constructors

        public ReadRange()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (ExcelRange == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ExcelRange)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var excelrange = ExcelRange.Get(context);

            string range = excelrange;



            // Create an Excel application object
            Application excelApp = new Application();

            Console.WriteLine("Excel Application Started");

            // Get the active workbook
            Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                Console.WriteLine("No open workbook found.");
                excelApp.Quit();
                return null;
            }
            Console.WriteLine("Workbook Detected");

            // Get the active worksheet
            Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

            // Read the data from the specified range
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.Range[range];

            object[,] data = (object[,])excelRange.Value;

            // Create a DataTable
            System.Data.DataTable dataTable = new System.Data.DataTable();

            // Get the dimensions of the data array
            int rowCount = data.GetLength(0);
            int columnCount = data.GetLength(1);

            // Add columns to the DataTable
            for (int col = 1; col <= columnCount; col++)
            {
                object headerValue = data[1, col];
                dataTable.Columns.Add(headerValue.ToString(), typeof(string));
            }

            // Add rows to the DataTable
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= columnCount; col++)
                {
                    object cellValue = data[row, col];
                    dataRow[col - 1] = cellValue.ToString();
                }
                dataTable.Rows.Add(dataRow);
            }

            // Clean up Excel objects
            Marshal.ReleaseComObject(excelrange);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);




            // Outputs
            return (ctx) => {
                DataOut.Set(ctx, null);
            };
        }

        #endregion
    }
}

