using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TurkishExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set Turkish culture for proper character handling
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            // Prompt the user to enter the file path
            // Console.WriteLine("Enter the file path of the Excel file:");
            // string filePath = Console.ReadLine();


            string filePath = @"C:\Users\hasan\Desktop\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa (YENİ)_995_03.30.2023_11.50.47.xlsx";





            // Create an Excel application object
            Application excelApp = null;
            try
            {
                excelApp = new Application();
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to create Excel application.");
                return;
            }

            // Open the workbook
            Workbook workbook = null;
            try
            {
                workbook = excelApp.Workbooks.Open(filePath);
            }
            catch (COMException)
            {
                Console.WriteLine("Failed to open the Excel file.");
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                return;
            }

            Console.WriteLine("Excel file opened successfully.");

            // Clean up Excel objects
            excelApp.Visible = true;
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
