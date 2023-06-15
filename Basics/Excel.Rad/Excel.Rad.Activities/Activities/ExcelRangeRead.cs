using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Excel.Rad.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace Excel.Rad.Activities
{
    [LocalizedDisplayName(nameof(Resources.ExcelRangeRead_DisplayName))]
    [LocalizedDescription(nameof(Resources.ExcelRangeRead_Description))]
    public class ExcelRangeRead : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ExcelRangeRead_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ExcelRangeRead_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ExcelRangeRead_Data_DisplayName))]
        [LocalizedDescription(nameof(Resources.ExcelRangeRead_Data_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> Data { get; set; }

        #endregion


        #region Constructors

        public ExcelRangeRead()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Range == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Range)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var range = Range.Get(context);
    
            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////

            // Outputs
            return (ctx) => {
                Data.Set(ctx, null);
            };
        }

        #endregion
    }
}

