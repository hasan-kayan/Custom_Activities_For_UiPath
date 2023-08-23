using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Text.RegularExpressions;
using System.Text;
using System.Globalization;

namespace ExcelReadApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;

            Console.WriteLine("Enter the path to the Excel file:");
            string filePath = Console.ReadLine();

            Console.WriteLine("Enter the range to read (e.g., A1:B5):");
            string range = Console.ReadLine();

            Console.WriteLine("Please enter the worksheet name you want to work on:");
            string worksheetName = Console.ReadLine();

            // Convert the file path and worksheet name strings to Unicode encoding
            byte[] filePathBytes = Encoding.Unicode.GetBytes(filePath);
            byte[] worksheetNameBytes = Encoding.Unicode.GetBytes(worksheetName);

            filePath = Encoding.Unicode.GetString(filePathBytes);
            worksheetName = Encoding.Unicode.GetString(worksheetNameBytes);

            Application app = new Application();
            app.Visible = false;

            Workbook existingWorkbook = null;
            Worksheet worksheet = null;

            try
            {
                existingWorkbook = app.Workbooks.Open(filePath);

                // Convert the worksheet name to uppercase and remove Turkish diacritics
                worksheetName = RemoveTurkishDiacritics(worksheetName.ToUpper());

                worksheet = GetWorksheetByName(existingWorkbook, worksheetName);

                if (worksheet != null)
                {
                    Range excelRange = worksheet.Range[range];
                    object[,] values = excelRange.Value;

                    int rowCount = values.GetLength(0);
                    int columnCount = values.GetLength(1);

                    Console.WriteLine($"Reading range: {range}");
                    Console.WriteLine();

                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int column = 1; column <= columnCount; column++)
                        {
                            object value = values[row, column];
                            Console.Write(value + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    Console.WriteLine($"Worksheet '{worksheetName}' not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (existingWorkbook != null)
                {
                    existingWorkbook.Close();
                    Marshal.ReleaseComObject(existingWorkbook);
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                worksheet = null;
                existingWorkbook = null;
                app = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.ReadLine();
        }

        static Worksheet GetWorksheetByName(Workbook workbook, string worksheetName)
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                string cleanWorksheetName = RemoveTurkishDiacritics(worksheet.Name.ToUpper());

                if (cleanWorksheetName == worksheetName)
                {
                    return worksheet;
                }
            }

            return null;
        }

        static string RemoveTurkishDiacritics(string input)
        {
            string normalizedString = input.Normalize(NormalizationForm.FormD);
            StringBuilder stringBuilder = new StringBuilder();

            foreach (char c in normalizedString)
            {
                UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}
