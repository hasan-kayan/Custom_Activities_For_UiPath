using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;

namespace LinkteraRobotics.Legacy.OpenExcelFile
{
        [Category("LinkteraRobotics.Legacy")]
        [DisplayName("Open Excel File")]
        [Description("This activity activates target Excel files.")]
        public class OpenExcelFileActivity : CodeActivity
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

            protected override void Execute(CodeActivityContext context)
            {
                string path = Path.Get(context);
                string sheetname = Sheetname.Get(context);

                Console.WriteLine("Linktera Robotics");

                // Excel Application starts
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                Console.WriteLine("Workbook detection...");
                Excel.Workbook workbook = excelApp.Workbooks.Open(path);

                // Find the Worksheet with the specified name
                Excel.Worksheet worksheet = null;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == sheetname)
                    {
                        worksheet = ws;
                        break;
                    }
                }

                // Check if the specified sheet exists, then activate it
                if (worksheet != null)
                {
                    worksheet.Activate();
                    Console.WriteLine("Sheet activated: " + worksheet.Name);
                }
                else
                {
                    Console.WriteLine("Sheet not found: " + sheetname);
                }
            
        }
    }
}