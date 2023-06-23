using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Excel.Range.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace LinkteraRobotics.Read.Excel.Range.Activities
{
    [LocalizedDisplayName(nameof(Resources.Read_DisplayName))]
    [LocalizedDescription(nameof(Resources.Read_Description))]
    public class Read : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.Read_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.Read_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.Read_Outputdata_DisplayName))]
        [LocalizedDescription(nameof(Resources.Read_Outputdata_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> Outputdata { get; set; }

        #endregion


        #region Constructors

        public Read()
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
                Outputdata.Set(ctx, null);
            };
        }

        #endregion
    }
}

