using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Big.Data.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace LinkteraRobotics.Read.Big.Data.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadLargeExcelData_Description))]
    public class ReadLargeExcelData : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_ExcelFilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_ExcelFilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> ExcelFilePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadLargeExcelData_OutputTable_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadLargeExcelData_OutputTable_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> OutputTable { get; set; }

        #endregion


        #region Constructors

        public ReadLargeExcelData()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));
            if (ExcelFilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(ExcelFilePath)));
            if (Range == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Range)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var sheetname = Sheetname.Get(context);
            var excelfilepath = ExcelFilePath.Get(context);
            var range = Range.Get(context);
    
            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            // Outputs
            return (ctx) => {
                OutputTable.Set(ctx, null);
            };
        }

        #endregion
    }
}

