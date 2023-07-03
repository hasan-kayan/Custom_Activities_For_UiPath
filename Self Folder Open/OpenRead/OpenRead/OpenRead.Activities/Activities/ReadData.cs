using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using OpenRead.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using ExcelDataReader;
using System.IO;
using System.ComponentModel;

namespace OpenRead.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadData_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadData_Description))]
    public class ReadData : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadData_FilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadData_FilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FilePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadData_RangeData_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadData_RangeData_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> RangeData { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadData_SheetName_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadData_SheetName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> SheetName { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadData_OutputData_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadData_OutputData_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> OutputData { get; set; }

        #endregion


        #region Constructors

        public ReadData()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (FilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FilePath)));
            if (RangeData == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(RangeData)));
            if (SheetName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(SheetName)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = FilePath.Get(context);
            var rangedata = RangeData.Get(context);
            var sheetname = SheetName.Get(context);

            using (var stream = File.Open(filepath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    var dataTable = dataSet.Tables[sheetname];
                    var dataView = new DataView(dataTable);
                    dataView.RowFilter = rangedata;


                    dataTable = OutputData;
                    // Outputs
                    return (ctx) =>
                    {
                        OutputData.Set(ctx, null);
                    };

                }

                #endregion
            }
        }
    }
}

