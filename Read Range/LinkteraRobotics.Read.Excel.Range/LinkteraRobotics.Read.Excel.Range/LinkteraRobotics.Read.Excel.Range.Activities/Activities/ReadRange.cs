using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Excel.Range.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;

using DataTable = System.Data.DataTable;
using System.Runtime.InteropServices;

namespace LinkteraRobotics.Read.Excel.Range.Activities
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

        [LocalizedDisplayName(nameof(Resources.ReadRange_Filepath_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Filepath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Filepath { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_Output_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Output_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> Output { get; set; }

        #endregion


        #region Constructors

        public ReadRange()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Filepath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Filepath)));
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));
            if (Range == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Range)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var sheetname = Sheetname.Get(context);
            var range = Range.Get(context);

            ///////////////////////////
            // Add execution logic HERE


            // Create an Excel application object
            Application excelApp = null;
            try
            {
                excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                throw new Exception("No active Excel instance found.");
            }

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
            Worksheet worksheet = workbook.Sheets[sheetname] as Worksheet;
            if (worksheet == null)
            {
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                throw new Exception($"Worksheet '{sheetname}' not found.");
            }

            // Read the range from Excel
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.Range[range];

            // Get the data from the range
            object[,] excelData = (object[,])excelRange.Value;

            // Convert the data to a DataTable
            DataTable dataTable = new DataTable();
            int rowCount = excelData.GetLength(0);
            int columnCount = excelData.GetLength(1);

            // Create columns in the DataTable
            for (int col = 1; col <= columnCount; col++)
            {
                string columnName = excelData[1, col]?.ToString() ?? $"Column{col}";
                dataTable.Columns.Add(columnName);
            }

            // Add rows to the DataTable
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= columnCount; col++)
                {
                    dataRow[col - 1] = excelData[row, col];
                }
                dataTable.Rows.Add(dataRow);
            }


            // Clean up Excel objects
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(worksheet);

            // If the workbook was opened in this execution, close it
            if (workbook != null && !workbook.FullName.Equals(filepath, StringComparison.OrdinalIgnoreCase))
            {
                workbook.Close();
            }
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);




            // Outputs
            return (ctx) => {
                Output.Set(ctx, dataTable);
            };
        }

        #endregion
    }
}

