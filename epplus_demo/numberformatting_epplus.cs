using OfficeOpenXml;
using System.Globalization;
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

            //Formatting column B to show integers
            sheet.Cells["B:B"].Style.Numberformat.Format = "0";

            //Formatting column C to show percent values with 2 decimal places
            sheet.Cells["C:C"].Style.Numberformat.Format = "0.00%";

            //Formatting column D to show values as Datetime
            sheet.Cells["D:D"].Style.Numberformat.Format = "dd/MM/yyyy HH:mm";

            //Formatting column E to show values as a Short Date. (Requires System.Globalization)
            sheet.Cells["E:E"].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

            //Saving the package
            package.SaveAs(new FileInfo(@"numberFormatting.xlsx"));

        }
    }
}
