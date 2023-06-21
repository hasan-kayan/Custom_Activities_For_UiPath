using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using Linktera.ReadRange.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Net.Mime;

namespace Linktera.ReadRange.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadExcelRange_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadExcelRange_Description))]
    public class ReadExcelRange : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadExcelRange_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadExcelRange_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadExcelRange_WorksheetName_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadExcelRange_WorksheetName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> WorksheetName { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadExcelRange_DataOutput_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadExcelRange_DataOutput_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> DataOutput { get; set; }

        #endregion


        #region Constructors

        public ReadExcelRange()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Range == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Range)));
            if (WorksheetName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(WorksheetName)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var range = Range.Get(context);
            var worksheetname = WorksheetName.Get(context);


            string Range = range;

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            // Prompt the user for the range to read
            Console.WriteLine("Enter the range to read (e.g., A1:C5):");
            

            // Attempt to find an open Excel instance
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                Console.WriteLine("No open Excel instance found. Error: " + ex.Message);
                throw;
            }

            // Get the active worksheet from the Excel instance
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;

            // Read the range of data and add it to a DataTable
            DataTable dataTable = new DataTable();
            Microsoft.Office.Interop.Excel.Range excelRange = activeSheet.Range[range];
            object[,] values = excelRange.Value;
            int rowCount = values.GetLength(0);
            int colCount = values.GetLength(1);

            // Add columns to the DataTable
            for (int col = 1; col <= colCount; col++)
            {
                object value = values[1, col];
                dataTable.Columns.Add(value.ToString());
            }

            // Add rows to the DataTable
            for (int row = 2; row <= rowCount; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= colCount; col++)
                {
                    object value = values[row, col];
                    dataRow[col - 1] = value;
                }
                dataTable.Rows.Add(dataRow);
            }

            // Close the Excel instance
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(activeSheet);
            Marshal.ReleaseComObject(excelApp);
            excelRange = null;
            activeSheet = null;
            excelApp = null;

            // Display the DataTable contents
            Console.WriteLine("\nData read from Excel:");
            foreach (DataRow dataRow in dataTable.Rows)
            {
                foreach (var item in dataRow.ItemArray)
                {
                    Console.Write(item.ToString() + "\t");
                }
            }

            // Outputs
            return (ctx) => {
                DataOutput.Set(ctx, null);
            };
        }

        #endregion
    }
}

