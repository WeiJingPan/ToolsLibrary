using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace MSExcelHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            var sw = new Stopwatch();
            sw.Start();
            var _xlsApp = new ApplicationClass();
            //打开已有的xls文件
            var xlsBook = _xlsApp.Workbooks.Open(@"D:\C_Spaces\ToolsLibrary\MSExcelHandler\MSExcelHandler\First.xlsx",
                Missing.Value, Missing.Value, Missing.Value, Missing.Value
                , Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //新建一个xls文件
            //xlsBook = _application.Workbooks.Add(Missing.Value);

            //指定要操作的Sheet，方式一：
            //var xlsSheet = (Worksheet)xlsBook.Sheets[1];
            //指定要操作的Sheet，方式二：
            var xlsSheet = (Worksheet) xlsBook.ActiveSheet;

            //指定单元格，读取数据，两种方法：
            //1
            var range1 = xlsSheet.get_Range("A1", Type.Missing);
            Console.WriteLine(range1.Value2);
            range1 = xlsSheet.get_Range("B1", Type.Missing);
            Console.WriteLine(range1.Value2);

            //2
            var range2 = (Range) xlsSheet.Cells[1, 3];
            Console.WriteLine(range2.Value2);

            var range3 = xlsSheet.get_Range("C1", Type.Missing);
            range3.Value2 = "Hello World";
            range3.Borders.Color = Color.FromArgb(123, 231, 32).ToArgb();
            range3.Font.Color = Color.Red.ToArgb();
            range3.Font.Size = 9;
            range3.Orientation = 90;//vertical
            range3.Columns.HorizontalAlignment = Constants.xlCenter;
            range3.Columns.VerticalAlignment = Constants.xlCenter;
            range3.Interior.Color = Color.FromArgb(192, 192, 192).ToArgb();
            range3.Columns.AutoFit(); //自动调整列宽度
            //在某个区域写入数据数组
            var matrixHeight = 20;
            var matrixWidth = 20;

            string[,] martix = new string[matrixHeight, matrixWidth];

            for (int i = 0; i < matrixHeight; i++)
                for (int j = 0; j < matrixWidth; j++)
                {
                    martix[i, j] = String.Format("{0}_{1}", i + 1, j + 1);
                }

            string startColName = GetColumnNameByIndex(0);

            string endColName = GetColumnNameByIndex(matrixWidth - 1);

            //取得某个区域，两种方法

            //之一：

            Range range4 = xlsSheet.get_Range("A1", Type.Missing);

            range4 = range4.get_Resize(matrixHeight, matrixWidth);

            //之二：

            //Range range4 = xlsSheet.get_Range(String.Format("{0}{1}", startColName, 1), String.Format("{0}{1}", endColName, martixHeight));

            range4.Value2 = martix;

            range4.Font.Color = Color.Red.ToArgb();

            range4.Font.Name = "Arial";

            range4.Font.Size = 9;

            range4.Columns.HorizontalAlignment = Constants.xlCenter;

            //设置column和row的宽度和颜色

            int columnIndex = 3;

            int rowIndex = 3;

            string colName = GetColumnNameByIndex(columnIndex);

            xlsSheet.get_Range(colName + rowIndex.ToString(), Type.Missing).Columns.ColumnWidth = 20;

            xlsSheet.get_Range(colName + rowIndex.ToString(), Type.Missing).Rows.RowHeight = 40;

            xlsSheet.get_Range(colName + rowIndex.ToString(), Type.Missing).Columns.Interior.Color = Color.Blue.ToArgb(); //单格颜色

            xlsSheet.get_Range(5 + ":" + 7, Type.Missing).Rows.Interior.Color = Color.Yellow.ToArgb(); //第5行到第7行的颜色

            //xlsSheet.get_Range("G : G", Type.Missing).Columns.Interior.Color=Color.Pink.ToArgb();//第n列的颜色如何设置？？

            //保存，关闭

            var xlxsPath = @"D:\C_Spaces\ToolsLibrary\MSExcelHandler\MSExcelHandler\Second.xlsx";

            if (File.Exists(xlxsPath))

            {
                File.Delete(xlxsPath);
            }

            xlsBook.SaveAs(xlxsPath, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            xlsBook.Close(false, Type.Missing, Type.Missing);

            _xlsApp.Quit();

            GC.Collect();

            Console.ReadKey();

            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            //Console.WriteLine("花费时间是：{0}", sw.ElapsedMilliseconds * 1000.0 / Stopwatch.Frequency);
            Console.WriteLine("花费时间是：{0}", ts2.TotalMilliseconds / 1000.0f);
            */
            
            //在某个区域写入数据数组
            var savePath = @"D:\C_Spaces\ToolsLibrary\MSExcelHandler\MSExcelHandler\Third.xlsx";
            // var matrixHeight = 20;
            // var matrixWidth = 20;
            // var martix = new string[matrixHeight, matrixWidth];
            //
            // for (var i = 0; i < matrixHeight; i++)
            //     for (var j = 0; j < matrixWidth; j++)
            //     {
            //         martix[i, j] = String.Format("{0}_{1}", i + 1, j + 1);
            //     }
            //
            // MSExcelManager.Inst.WriteToExcel(savePath, martix, matrixHeight, matrixWidth);

            var list_range = MSExcelManager.Inst.ReadingExcel(savePath);
            foreach (var row in list_range.Rows)
            {
                var curRow = (DataRow)row;
                // foreach (var column in curRow)
                // {
                //     Console.WriteLine("{0}\t", column);
                // }
                
                Console.WriteLine("\n");
            }

        }
        
        //将column index转化为字母，至多两位
        public static string GetColumnNameByIndex(int index)

        {
            string[] alphabet = new string[]
            {
                "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
                "U", "V", "W", "X", "Y", "Z"
            };

            string result = "";

            int temp = index / 26;

            int temp2 = index % 26 + 1;

            if (temp > 0)

            {
                result += alphabet[temp];
            }

            result += alphabet[temp2];

            return result;
        }
    }
}