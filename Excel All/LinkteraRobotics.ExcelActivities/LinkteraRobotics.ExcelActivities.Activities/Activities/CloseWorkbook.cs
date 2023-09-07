using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.ExcelActivities.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace LinkteraRobotics.ExcelActivities.Activities
{
    [LocalizedDisplayName(nameof(Resources.CloseWorkbook_DisplayName))]
    [LocalizedDescription(nameof(Resources.CloseWorkbook_Description))]
    public class CloseWorkbook : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CloseWorkbook_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.CloseWorkbook_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        [LocalizedDisplayName(nameof(Resources.CloseWorkbook_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.CloseWorkbook_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        #endregion


        #region Constructors

        public CloseWorkbook()
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

            // Outputs
            return (ctx) => {
            };
        }

        #endregion
    }
}

