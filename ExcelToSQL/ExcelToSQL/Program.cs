//// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using System;
using System.Text;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Add this line

        using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\Users\DELL\Documents\Excel\123.xlsx")))
        {
            StringBuilder sb = new StringBuilder();
            var worksheet = package.Workbook.Worksheets[0]; // assuming you want the first worksheet

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            for (int row = 2; row <= rowCount; row++)
            {
                sb.Append("INSERT INTO YourTableName (");

                for (int col = 1; col <= colCount; col++)
                {
                    sb.Append(worksheet.Cells[1, col].Value.ToString());

                    if (col < colCount)
                    {
                        sb.Append(", ");
                    }
                }

                sb.Append(") VALUES (");

                for (int col = 1; col <= colCount; col++)
                {
                    sb.Append("'" + worksheet.Cells[row, col].Value.ToString() + "'");

                    if (col < colCount)
                    {
                        sb.Append(", ");
                    }
                }

                sb.AppendLine(");");
            }

            Console.WriteLine(sb.ToString());
        }
    }
}
