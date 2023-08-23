using System;
using System.Data;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;


using DataTable = System.Data.DataTable;
namespace ExcelConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Inputs
            string filepath = @"C:\Users\hasan\Desktop\Büyük Excel\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa (YENİ)_995_03.30.2023_11.50.47.xlsx";
            string targetsheet = "TC Hazine ve Maliye Bakanlığı y";
            string targetrange = "A1:B5";

            // Create an Excel application object
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Check if the workbook is already open
            Workbook workbook = null;
            foreach (Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.Equals(filepath, StringComparison.OrdinalIgnoreCase))
                {
                    workbook = wb;
                    break;
                }
            }

            // If the workbook is not open, open it
            if (workbook == null)
            {
                workbook = excelApp.Workbooks.Open(filepath);
            }

            // Get the worksheet
            Worksheet worksheet = workbook.Sheets[targetsheet] as Worksheet;
            if (worksheet == null)
            {
                workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                throw new Exception($"Worksheet '{targetsheet}' not found.");
            }

            // Add a new worksheet
            Worksheet newWorksheet = workbook.Sheets.Add(After: workbook.ActiveSheet) as Worksheet;
            Console.WriteLine("New worksheet created");

            try
            {
                Console.WriteLine("Copying data to new worksheet");
                // Select and copy all cells from the source worksheet
                Range cells = worksheet.Cells;
                cells.Select();
                cells.Copy();
                



                // Paste the copied cells to the new worksheet
                Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);

                Console.WriteLine("Kopyalama ve yapıştırma işlemi tamamlandı.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata: " + ex.Message);
            }

            try
            {
                // Read data from the new worksheet
                Range excelRange;
                if (string.IsNullOrWhiteSpace(targetrange))
                {
                    excelRange = newWorksheet.UsedRange;
                }
                else
                {
                    excelRange = newWorksheet.Range[targetrange];
                }

                // Read the data
                object[,] excelData = (object[,])excelRange.Value;

                // Create a DataTable and transfer the data
                DataTable dataTable = new DataTable();
                int rowCount = excelData.GetLength(0);
                int columnCount = excelData.GetLength(1);

                for (int col = 1; col <= columnCount; col++)
                {
                    string columnName = excelData[1, col]?.ToString() ?? $"Column{col}";
                    dataTable.Columns.Add(columnName);
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        dataRow[col - 1] = excelData[row, col];
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // Clear Excel objects
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorksheet);

                // If the workbook was opened in this session, close it
                if (workbook != null && !workbook.FullName.Equals(filepath, StringComparison.OrdinalIgnoreCase))
                {
                    workbook.Close();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // You can perform the desired operations using the DataTable
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata: " + ex.Message);
            }

            Console.ReadLine();
        }
    }
}
