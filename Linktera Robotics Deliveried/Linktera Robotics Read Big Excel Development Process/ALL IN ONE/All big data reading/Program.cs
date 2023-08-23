using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace All_big_data_reading
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\Users\hasan\Desktop\Büyük Excel\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa_Robotik Süreç (1).xlsx";
            string sheetName = "TC Hazine ve Maliye Bakanlığı y";
            string range = "";
            range = null;
            Console.WriteLine("New Application Starting...");

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Console.WriteLine("Workbook detection...");
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = null;
            Excel.Worksheet newWorksheet = null;

            try
            {
                worksheet = workbook.Sheets[sheetName];

                Console.WriteLine("Copying Cells");
                Excel.Range cells = worksheet.Cells;
                cells.Copy();
                Console.WriteLine("Cells Copied");

                Console.WriteLine("Creating New Sheet");
                newWorksheet = workbook.Sheets.Add(After: workbook.ActiveSheet);
                newWorksheet.Name = "Copied";
                Console.WriteLine("New Sheet named 'Copied' Created");

                Excel.Range pasteRange = newWorksheet.Cells;
                pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    SkipBlanks: false, Transpose: false);
                Console.WriteLine("Paste Special Completed");

                if (newWorksheet != null)
                {
                    Console.WriteLine("Data Table Creating, data reading...");
                    Excel.Range excelRange;
                    if (string.IsNullOrWhiteSpace(range))
                    {
                        Console.WriteLine("All data reading");
                        excelRange = newWorksheet.UsedRange;
                    }
                    else
                    {
                        Console.WriteLine("The range : " + range + "is reading");
                        excelRange = newWorksheet.Range[range];
                    }

                    int rowCount = excelRange.Rows.Count;
                    int columnCount = excelRange.Columns.Count;

                    DataTable dataTable = new DataTable();

                    // Read column names
                    for (int col = 1; col <= columnCount; col++)
                    {
                        string columnName = excelRange.Cells[1, col]?.Value?.ToString() ?? $"Column{col}";
                        dataTable.Columns.Add(columnName);
                    }
                    Console.WriteLine("Column names read");

                    // Read data in smaller chunks (rows)
                    const int batchSize = 1000;
                    int remainingRows = rowCount - 1;
                    int startRow = 2;

                    while (remainingRows > 0)
                    {
                        int rowsToRead = Math.Min(batchSize, remainingRows);
                        int endRow = startRow + rowsToRead - 1;

                        Excel.Range dataRange = excelRange.Range[excelRange.Cells[startRow, 1], excelRange.Cells[endRow, columnCount]];
                        object[,] excelData = (object[,])dataRange.Value;
                        Console.WriteLine($"Read rows {startRow} to {endRow}");

                        // Transfer data to the DataTable
                        for (int row = 1; row <= rowsToRead; row++)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            for (int col = 1; col <= columnCount; col++)
                            {
                                dataRow[col - 1] = excelData[row, col];
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                        Console.WriteLine($"Transferred rows {startRow} to {endRow}");

                        startRow += batchSize;
                        remainingRows -= rowsToRead;
                    }

                    Console.WriteLine("Data Table Content:");
                    foreach (DataRow row in dataTable.Rows)
                    {
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            Console.Write(row[col] + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    Console.WriteLine("New worksheet is null. Cannot proceed with data reading.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }
            finally
            {
                newWorksheet?.Activate();
                workbook.Close(false); // False states that dont save changes !!!!!!!!!
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.ReadKey();
        }
    }
}
