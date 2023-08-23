using System;
using Microsoft.Office.Interop.Excel;
using System.IO;

class Program
{
    static void Main()
    {
        string excelFilePath = @"C:\Users\hasan\Desktop\Büyük Excel\Halkbank\TC Hazine ve Maliye Bakanlığı yazısı - İhracat bedelleri+IBKB_V2_Exa (YENİ)_995_03.30.2023_11.50.47.xlsx";
        string sheetName = "TC Hazine ve Maliye Bakanlığı y";
        int rowsPerIteration = 40000;

        string startColumnName = "A";
        int startColumnValue = 1;

        string endColumnName = "BJ";
        int endColumnValue = 20000;

        int endColumnValueFirst = startColumnValue + 4000;

        float rowCount = endColumnValue - startColumnValue;
        int loopCount = (int)rowCount;

        try
        {
            Console.WriteLine("Excel dosyası açılıyor...");
            Application excelApp = new Application();
            excelApp.Visible = true;

            Workbook workbook = null;
            Worksheet worksheet = null;

            // Excel dosyasının zaten açık olup olmadığını kontrol et
            foreach (Workbook openWorkbook in excelApp.Workbooks)
            {
                if (openWorkbook.FullName == excelFilePath)
                {
                    workbook = openWorkbook;
                    worksheet = workbook.Sheets[sheetName];
                    break;
                }
            }

            if (workbook == null)
            {
                Console.WriteLine("Excel dosyası zaten açık değil. Dosya açılıyor...");
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[sheetName];
            }

            // Yeni bir Excel dosyası oluştur
            string newExcelFilePath = Path.Combine(Path.GetDirectoryName(excelFilePath), "copy.xlsx");
            Workbook newWorkbook = excelApp.Workbooks.Add();
            Worksheet newWorksheet = newWorkbook.Sheets[1];

            for (int i = 1; i <= loopCount; i++)
            {
                string startRange = startColumnName + startColumnValue + ":" + endColumnName + (int)endColumnValueFirst;
                

                Console.WriteLine(startRange + " okunuyor...");

                // Verileri oku
                Range range = worksheet.Range[startRange];
                object[,] values = (object[,])range.Value;

                // Verileri copy.xlsx dosyasına yaz
                Range newRange = newWorksheet.Range[startColumnName + (startColumnValue + 1) + ":" + endColumnName + (endColumnValueFirst + 1)];
                newRange.Value = values;

                // Veri tutma değerlerini sıfırla
                values = null;
                range = null;

                startColumnValue = endColumnValueFirst + 1;
                endColumnValueFirst = endColumnValueFirst + 4000;
            }

            // Excel dosyalarını kapat ve kaynakları serbest bırak
            workbook.Close();
            newWorkbook.Save();
            newWorkbook.Close();
            excelApp.Quit();

            Console.WriteLine("İşlem tamamlandı. Yeni Excel dosyası: " + newExcelFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Hata: " + ex.Message);
        }
        Console.ReadKey();
    }
}
