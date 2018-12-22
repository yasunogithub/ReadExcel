using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab

namespace ReadExcel
{
    class Program
    {
        // 参考 https://coderwall.com/p/app3ya/read-excel-file-in-c
        // 参照=>COM=>Microsoft Excel 16 objectを追加する
        static void Main(string[] args)
        {
            
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open($@"{Directory.GetCurrentDirectory()}\sandbox_1.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            var last = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow, lastColumn;
            lastRow= last.Row;
            lastColumn = last.Column;
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= lastRow; i++)
            {
                for (int j = 1; j <= lastColumn; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    //add useful things here!   
                }
            }
            Console.ReadKey();
            //close and release
            xlWorkbook.Close();
            //quit and release
            xlApp.Quit();
        }
    }
}
