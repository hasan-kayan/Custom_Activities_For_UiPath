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
    [DisplayName("Close Excel File")]
    [Description("This activity closes the target Excel file.")]
    public class OpenExcelFileActivity : CodeActivity
    {
        [Category("Input")]
        [DisplayName("Path")]
        [Description("Enter the path of the Excel file.")]
        [RequiredArgument]
        public InArgument<string> Path { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string targetFilePath = Path.Get(context);

            Console.WriteLine("Linktera Robotics");

            Excel.Application excelApp = null;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                foreach (Excel.Workbook workbook in excelApp.Workbooks)
                {
                    if (workbook.FullName == targetFilePath)
                    {
                        workbook.Close(false); // Close without saving changes
                        Marshal.ReleaseComObject(workbook);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }
}