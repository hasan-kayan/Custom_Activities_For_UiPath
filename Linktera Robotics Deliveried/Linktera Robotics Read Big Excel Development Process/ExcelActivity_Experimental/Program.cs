using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel; // .NET version excel usage microsoft lib
using Microsoft.SqlServer.Server;
// For Excel applications C# has  Microsoft.Office.Interop.Excel namespace, namespace using interface architecture
// Namespace includes Application class which represents Excel Application


namespace ExcelReadActivity
{
    public class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();  // Creating new instance of the ' Application Class '. If you have to declare again remember you it will give error 
            // Go on the error and choose namespace 
            app.Visible = true; // Set visible property of the "Application" object to "true"


            Workbook sampleWorkbook = app.Workbooks.Add(); // Creating Workbook, added application to manage


            Console.WriteLine("Enter Path:");
            string pat = Console.ReadLine();

            Console.WriteLine("Enter sheet:");
            string sheet = Console.ReadLine();



            // So far we have created an empty Workbook, app is representing Excel file so we have created a new Excel file 

            Workbook exsistingWorkbook = app.Workbooks.Open(pat); // Thats how you can open an existing workbook 

            // So far we have created an empty workbook and opened an exsiting one 

            // Declear worksheet object

            Worksheet worksheet = sampleWorkbook.Worksheets["sheet1"]; 

            // change value of one cell 

            worksheet.Range["A1"].Value = "Deneme Outt";

            double[] SalesDate = { 4.3, 4, 21, 324, 17 };

            for (int i = 0; i < SalesDate.Length; i++)
            {
                worksheet.Range["A" + (2+i)].Value= SalesDate[i];
            }

            // Write same data to multiple cells 






        }
    }
}
// C:\Users\hasan\Desktop\excel applications try\deneme.xlsx\