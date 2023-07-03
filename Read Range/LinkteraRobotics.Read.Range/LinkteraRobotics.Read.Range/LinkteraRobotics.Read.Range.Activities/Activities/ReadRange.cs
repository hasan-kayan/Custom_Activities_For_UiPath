using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Range.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using System.ComponentModel;

using DataTable = System.Data.DataTable;

// Your code here

namespace LinkteraRobotics.Read.Range.Activities
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

        [LocalizedDisplayName(nameof(Resources.ReadRange_Datarange_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Datarange_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Datarange { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_ReadedData_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_ReadedData_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> ReadedData { get; set; }

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
            if (Datarange == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Datarange)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var sheetname = Sheetname.Get(context);
            var datarange = Datarange.Get(context);



      





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

