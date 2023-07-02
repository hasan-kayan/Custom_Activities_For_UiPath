using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.Read.Range.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace LinkteraRobotics.Read.Range.Activities
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

        [LocalizedDisplayName(nameof(Resources.ReadExcelRange_SheetName_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadExcelRange_SheetName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> SheetName { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadExcelRange_DataRange_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadExcelRange_DataRange_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> DataRange { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadExcelRange_ReadedData_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadExcelRange_ReadedData_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> ReadedData { get; set; }

        #endregion


        #region Constructors

        public ReadExcelRange()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (SheetName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(SheetName)));
            if (DataRange == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(DataRange)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var sheetname = SheetName.Get(context);
            var datarange = DataRange.Get(context);
    
            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            // Outputs
            return (ctx) => {
                ReadedData.Set(ctx, null);
            };
        }

        #endregion
    }
}

