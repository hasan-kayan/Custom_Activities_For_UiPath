using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using Linktera.Robotics.Excel.Basics.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace Linktera.Robotics.Excel.Basics.Activities
{
    [LocalizedDisplayName(nameof(Resources.OpenExcelScope_DisplayName))]
    [LocalizedDescription(nameof(Resources.OpenExcelScope_Description))]
    public class OpenExcelScope : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.OpenExcelScope_FilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.OpenExcelScope_FilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FilePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.OpenExcelScope_WorksheetName_DisplayName))]
        [LocalizedDescription(nameof(Resources.OpenExcelScope_WorksheetName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> WorksheetName { get; set; }

        #endregion


        #region Constructors

        public OpenExcelScope()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (FilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FilePath)));
            if (WorksheetName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(WorksheetName)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = FilePath.Get(context);
            var worksheetname = WorksheetName.Get(context);
    
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

