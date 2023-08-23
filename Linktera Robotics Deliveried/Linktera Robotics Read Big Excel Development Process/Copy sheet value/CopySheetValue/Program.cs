using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CopySheetValue
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Inputs
            Console.WriteLine("Enter the file path: ");
            string filepath = Console.ReadLine();
            Console.WriteLine("Enter the source sheet name: ");
            string sheetname = Console.ReadLine();
            Console.WriteLine("Enter the target sheet name: ");
            string sheetnamepaste = Console.ReadLine();

            // Excel application object
            var excelApp = new Application();
            Workbook workbook = null;

            try
            {
                // Check if the Excel file is already open
                try
                {
                    workbook = excelApp.Workbooks.get_Item(filepath);
                }
                catch (Exception)
                {
                    // File is not open, open it
                    workbook = excelApp.Workbooks.Open(filepath);
                }

                Worksheet sourceSheet = null;
                foreach (Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == sheetname)
                    {
                        sourceSheet = sheet;
                        break;
                    }
                }

                if (sourceSheet != null)
                {
                    // Add new target sheet
                    Worksheet targetSheet = (Worksheet)workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                    targetSheet.Name = sheetnamepaste;

                    // Copy source sheet data
                    sourceSheet.Cells.Copy();

                    // Paste values to target sheet
                    targetSheet.Select();
                    targetSheet.PasteSpecial();

                    Console.WriteLine("Sheet copied and pasted successfully.");
                }
                else
                {
                    Console.WriteLine("Source sheet not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error processing Excel file. " + ex.Message);
            }
            finally
            {
                // Close Excel application
                if (workbook != null)
                {
                    workbook.Close(SaveChanges: true);
                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                workbook = null;
                excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
