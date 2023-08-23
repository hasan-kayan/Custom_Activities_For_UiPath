using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.V4Legacy.ReadBigData.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using ExcelDataReader;
using System.IO;
using System.Text;


namespace LinkteraRobotics.V4Legacy.ReadBigData.Activities
{
    [LocalizedDisplayName(nameof(Resources.ReadBigData_DisplayName))]
    [LocalizedDescription(nameof(Resources.ReadBigData_Description))]
    public class ReadBigData : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_Path_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_Path_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Path { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_Sheetname_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_Sheetname_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Sheetname { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadBigData_Outdata_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadBigData_Outdata_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> Outdata { get; set; }

        #endregion


        #region Constructors

        public ReadBigData()
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
             DataTable dataTable = new DataTable();

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            try
            {


                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Read the header row and create columns in the DataTable
                        if (reader.Read())
                        {
                            for (int col = 0; col < reader.FieldCount; col++)
                            {
                                var columnName = reader.GetString(col);
                                dataTable.Columns.Add(columnName);
                            }
                        }

                        // Skip the header row and read data rows
                        while (reader.Read())
                        {
                            var rowData = new object[reader.FieldCount];
                            for (int col = 0; col < reader.FieldCount; col++)
                            {
                                rowData[col] = reader.GetValue(col);
                            }
                            dataTable.Rows.Add(rowData);
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                // Handle the exception if needed.
                Console.WriteLine($"Error occurred: {ex.Message}");
            }



            // Outputs
            return (ctx) => {
                Outdata.Set(ctx, dataTable);
            };
        }

        #endregion
    }
}

