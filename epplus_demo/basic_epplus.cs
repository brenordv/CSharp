using OfficeOpenXml;
using System.IO;

namespace epplus_basic
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creating empty Excel Package...
            var package = new ExcelPackage();

            //Referencing the Workbook...
            var workbook = package.Workbook;

            //Adding a new sheet named "SheetName"...
            var sheet = workbook.Worksheets.Add("SheetName");

            //Accessing cell A1 using row/col numbers.
            //Hint: EPPlus references cells using a base 1 system.
            var row = 1;
            var col = 1;

            //Setting the value of A1 to "Ping?"
            sheet.Cells[row, col].Value = "Ping?";

            //Setting the value of B1 to "Pong!"
            sheet.Cells["B1"].Value = "Pong!";

            //Setting the value of A3 to "Hey!"
            sheet.Cells["A3"].Value = "Hey!";
            //...and immediately changing it's value back to empty (null)... 
            sheet.Cells["A3"].Value = null;


            //Setting cells A1:E1 to bold.
            sheet.Cells["A1:E1"].Style.Font.Bold = true;
            

            //Saving the reference to cell A3 to a variable...
            var celA3 = sheet.Cells["A3"];

            //Saving the reference for the range A1:E1 to a variable...
            var rangeA1E1 = sheet.Cells["A1:E1"];

            //Saving the file. Passing the filename...
            package.SaveAs(new FileInfo(@"sampleEpplus.xlsx"));


            //EXTRA

            //Creating a new file and already setting a filename
            var blankPackage = new ExcelPackage(new FileInfo(@"blankSample.xlsx"));

            //Adding a sheet, so we can save it.
            blankPackage.Workbook.Worksheets.Add("blankSheet");

            //Saving the new file using the original filename.
            blankPackage.Save();

            //Creating a copy of that file.
            blankPackage.SaveAs(new FileInfo(@"blankSampleCopy.xlsx"));

        }
    }
}
