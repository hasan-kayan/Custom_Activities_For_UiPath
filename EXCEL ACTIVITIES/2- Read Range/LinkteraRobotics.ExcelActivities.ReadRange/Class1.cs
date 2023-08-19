using System;
using System.Activities;
using System.ComponentModel;
using OfficeOpenXml;
using OfficeOpenXml.Core.ExcelPackage;

namespace LinkteraRobotics.ExcelActivities.ReadRange
{
    [Category("Linktera Robotics.Excel Activities")]
    [DisplayName("Read Range")]
    [Description("This activity aims to read specific Excel range.")]
    public class OpenWorkbook : CodeActivity
    {
        [Category("Input")]
        [DisplayName("Path")]
        [Description("Enter the path of the Excel file.")]
        [RequiredArgument]
        public InArgument<string> Path { get; set; }

        [Category("Input")]
        [DisplayName("Sheet Name")]
        [Description("Enter the sheet name to open.")]
        [RequiredArgument]
        public InArgument<string> Sheetname { get; set; }

        [Category("Input")]
        [DisplayName("Range")]
        [Description("Enter the range to read.")]
        [RequiredArgument]
        public InArgument<string> Range { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string filePath = Path.Get(context);
            string sheetName = Sheetname.Get(context);
            string range = Range.Get(context);

            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"Sheet '{sheetName}' not found.");
                    return;
                }

                var selectedRange = worksheet.Cells[range];
                if (selectedRange == null)
                {
                    Console.WriteLine($"Range '{range}' not found in sheet '{sheetName}'.");
                    return;
                }

                foreach (var cell in selectedRange)
                {
                    Console.WriteLine($"Cell [{cell.Start.Row},{cell.Start.Column}]: {cell.Text}");
                }
            }
        }
    }
}
