using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using LinkteraRobotics.Copy.Sheet.Value.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace LinkteraRobotics.Copy.Sheet.Value.Activities
{
    [LocalizedDisplayName(nameof(Resources.CopySheetValue_DisplayName))]
    [LocalizedDescription(nameof(Resources.CopySheetValue_Description))]
    public class CopySheetValue : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheetValue_Filepath_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheetValue_Filepath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Filepath { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheetValue_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheetValue_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        [LocalizedDisplayName(nameof(Resources.CopySheetValue_SheetnamePaste_DisplayName))]
        [LocalizedDescription(nameof(Resources.CopySheetValue_SheetnamePaste_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> SheetnamePaste { get; set; }

        #endregion


        #region Constructors

        public CopySheetValue()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (Filepath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Filepath)));
            if (Sheetname == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Sheetname)));
            if (SheetnamePaste == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(SheetnamePaste)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var sheetname = Sheetname.Get(context);
            var sheetnamepaste = SheetnamePaste.Get(context);
    
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

