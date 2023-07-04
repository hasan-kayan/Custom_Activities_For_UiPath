using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using LinkteraRobotics.Read.Excel.Range.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office.Interop.Excel;

using DataTable = System.Data.DataTable;



namespace LinkteraRobotics.Read.Excel.Range.Activities
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

        [LocalizedDisplayName(nameof(Resources.ReadRange_Range_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Range_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Range { get; set; }

        [LocalizedDisplayName(nameof(Resources.ReadRange_Output_DisplayName))]
        [LocalizedDescription(nameof(Resources.ReadRange_Output_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<DataTable> Output { get; set; }

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
            if (Range == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(Range)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = Filepath.Get(context);
            var sheetname = Sheetname.Get(context);
            var range = Range.Get(context);

            ///////////////////////////
            // Add execution logic HERE


            // Excel uygulamasýný baþlat
            Application excelApp = new Application();

            // Excel dosyasýný aç
            Workbook workbook = excelApp.Workbooks.Open(filepath);

            // Ýlgili sayfayý seç
            Worksheet worksheet = (Worksheet)workbook.Sheets[sheetname];

            // Belirtilen aralýðý al
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.Range[range];



            // Verileri bir object dizisine aktar
            object[,] valueArray = (object[,])excelRange.Value;

            // DataTable oluþtur
            DataTable dataTable = new DataTable();



            // Sütun baþlýklarýný ekle
            for (int col = 1; col <= excelRange.Columns.Count; col++)
            {
                dataTable.Columns.Add(valueArray[1, col].ToString());
            }

            // Verileri DataTable'e aktar
            for (int row = 2; row <= excelRange.Rows.Count; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= excelRange.Columns.Count; col++)
                {
                    dataRow[col - 1] = valueArray[row, col];
                }
                dataTable.Rows.Add(dataRow);
            }

            // Excel uygulamasýný kapat
            excelApp.Quit();

            // DataTable'i kullanabilirsiniz
            // Örneðin, verileri yazdýrabilirsiniz
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (DataColumn col in dataTable.Columns)
                {
                    Console.Write(row[col] + "\t");
                }
                Console.WriteLine();
            }
            // Kaynaklarý serbest býrak
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // Bellekteki nesneleri temizle
            worksheet = null;
            workbook = null;
            excelApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Output ile ara datatable ý eþitle
            Output.Set(context, dataTable);





            // Outputs
            return (ctx) => {
                Output.Set(ctx, null);
            };
        }

        #endregion
    }
}

