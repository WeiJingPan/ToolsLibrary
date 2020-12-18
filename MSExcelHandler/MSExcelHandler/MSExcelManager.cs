/****** Author ******
* Created By : PWJ
* Date:2020-12-18
*/

using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace MSExcelHandler
{
    public class MSExcelManager
    {
        private static MSExcelManager _inst;
        public static MSExcelManager Inst { get => _inst??= new MSExcelManager(); }

        private Application _xlsApp;

        private string[] alphabet;
        
        public MSExcelManager()
        {
            _xlsApp??= new ApplicationClass();
            alphabet??= new[]
            {
                "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
                "U", "V", "W", "X", "Y", "Z"
            };
        }

        public System.Data.DataTable ReadingExcel(string path)
        {
            if (File.Exists(path))
            {
                
                var xlsBook = _xlsApp.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value);
                var xlsSheet = (Worksheet)xlsBook.Sheets[1];
                var curRange = xlsSheet.get_Range("A1", Type.Missing);
                curRange.get_Resize(xlsSheet.UsedRange.Rows.Count, xlsSheet.UsedRange.Columns.Count);
                xlsBook.Close(false, Type.Missing, Type.Missing);
                return (System.Data.DataTable)curRange.Value2;
            }
            Console.WriteLine("{0} 资源并不存在！", path);
            return null;
        }

        public void TestReadExcel(string path)
        {
            if (File.Exists(path))
            {
                
                var xlsBook = _xlsApp.Workbooks.Open(path, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value);
                var xlsSheet = (Worksheet)xlsBook.Sheets[1];
                var rowsCount = xlsSheet.UsedRange.Rows.Count;
                var columnCount = xlsSheet.UsedRange.Columns.Count;
                Console.WriteLine("rowsCount = {0}, columnCount = {1}", rowsCount, columnCount);
                
                xlsBook.Close(false, Type.Missing, Type.Missing);
                return;
            }
            Console.WriteLine("{0} 资源并不存在！", path);
        }

        public void WriteToExcel(string path, string[,] content, int matrixHeight, int matrixWidth)
        {
            
            if(File.Exists(path)) File.Delete(path);

            string startColName = GetColumnNameByIndex(0);
            string endColName = GetColumnNameByIndex(matrixWidth - 1);
            
            var xlsBook = _xlsApp.Workbooks.Add(Missing.Value);
            var xlsSheet = (Worksheet)xlsBook.ActiveSheet;
            xlsSheet.Name = "First_Sheet";
            var curRange = xlsSheet.get_Range("A1", Missing.Value);
            curRange = curRange.get_Resize(matrixHeight, matrixWidth);
            //curRange = xlsSheet.get_Range(String.Format("{0}{1}", startColName, 1), String.Format("{0}{1}", endColName, matrixHeight));
            curRange.Value2 = content;
            curRange.Font.Name = "Arial";
            curRange.Borders.Color = Color.FromArgb(123, 231, 32).ToArgb();
            curRange.Font.Color = Color.Red.ToArgb();
            curRange.Font.Size = 9;
            curRange.Orientation = 90;//vertical
            curRange.Columns.HorizontalAlignment = Constants.xlCenter;
            curRange.Columns.VerticalAlignment = Constants.xlCenter;
            curRange.Interior.Color = Color.FromArgb(192, 192, 192).ToArgb();
            curRange.Columns.AutoFit(); //自动调整列宽度

            xlsBook.SaveAs(path, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            xlsBook.Close(false, Type.Missing, Type.Missing);

            _xlsApp.Quit();

            GC.Collect();

            Console.WriteLine("生成表格成功！");
            
        }

        public string GetColumnNameByIndex(int index)
        {
            var result = "";
            var temp = index / 26;
            var temp2 = index % 26 + 1;
            if (temp > 0) result += alphabet[temp];
            result += alphabet[temp2];
            return result;
        }

    }
}