using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Data;
using System.IO;
using ExcelDataReader;

namespace LinkteraRobotics.Legacy.ReadBigData
{

    //Note that these attributes are localized so you need to localize this attribute for Studio languages other than English

    //Dots allow for hierarchy. App Integration.Excel is where Excel activities are.
    [Category("LinkteraRobotics.Legacy")]
    [DisplayName("Read Big Data")]
    [Description("This activty is based on to read big Excel data and keep it into a data table. Specially this activty aims to Windows-Legacy Copmatible.  ")]
    public class ReadBigDataLegacy : CodeActivity
    {
        [Category("Input")]
        [DisplayName("Path")]
        [Description("Enter the path of the Excel file.")]
        [RequiredArgument]
        public InArgument<string> Path { get; set; }

        [Category("Input")]
        [DisplayName("Sheetname")]
        [Description("Enter the name of the sheet to read from.")]
        [RequiredArgument]
        public InArgument<string> Sheetname { get; set; }

        [Category("Output")]
        [DisplayName("Outdata")]
        [Description("The DataTable containing the read data.")]
        public OutArgument<DataTable> Outdata { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var path = Path.Get(context);
            var sheetname = Sheetname.Get(context);

            var dataTable = new DataTable();

            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        if (reader.Read())
                        {
                            for (int col = 0; col < reader.FieldCount; col++)
                            {
                                var columnName = reader.GetString(col);
                                dataTable.Columns.Add(columnName);
                            }
                        }

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
                Console.WriteLine($"Error occurred: {ex.Message}");
            }

            Outdata.Set(context, dataTable);
        }
    }
    }
