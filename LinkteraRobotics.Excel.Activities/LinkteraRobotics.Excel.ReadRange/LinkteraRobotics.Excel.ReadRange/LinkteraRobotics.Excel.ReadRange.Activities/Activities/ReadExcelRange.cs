using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Excel.ReadRange.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel; 

namespace LinkteraRobotics.Excel.ReadRange.Activities
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
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                Console.WriteLine("No active Excel instance found.");
                throw;
            }


            // Outputs
            return (ctx) => {
                DataOutput.Set(ctx, null);
            };
        }

        #endregion
    }
}

