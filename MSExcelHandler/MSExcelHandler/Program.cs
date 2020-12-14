using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace MSExcelHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            var sw = new Stopwatch();
            sw.Start();
            var _xlsApp = new ApplicationClass();
            //打开已有的xls文件
            var xlsBook = _xlsApp.Workbooks.Open(@"C:\Users\v_wjjwpan\v_wjjwpan的同步盘\TestCSharp\TestCSharp\First.xlsx", Missing.Value, Missing.Value, Missing.Value, Missing.Value
                , Missing.Value, Missing.Value, Missing.Value,Missing.Value,Missing.Value);
            //新建一个xls文件
            //xlsBook = _application.Workbooks.Add(Missing.Value);
            
            //指定要操作的Sheet，方式一：
            //var xlsSheet = (Worksheet)xlsBook.Sheets[1];
            //指定要操作的Sheet，方式二：
            var xlsSheet = (Worksheet)xlsBook.ActiveSheet;
            
            //指定单元格，读取数据，两种方法：
            //1
            var range1 = xlsSheet.get_Range("A1", Type.Missing);
            Console.WriteLine(range1.Value2);
            range1 = xlsSheet.get_Range("B1", Type.Missing);
            Console.WriteLine(range1.Value2);
            
            //2
            var range2 = (Range)xlsSheet.Cells[1, 3];
            Console.WriteLine(range2.Value2);

            var range3 = xlsSheet.get_Range("C1", Type.Missing);
            range3.Value2 = "Hello World";
            range3.Borders.Color = Color.FromArgb(123, 231, 32).ToArgb();
            range3.Font.Color = Color.Red.ToArgb();
            range3.Font.Size = 9;
            //range3.Orientation = 90;//vertical
            //range3.Columns.HorizontalAlignment = Excel.Constants.xlCenter;
            // range3.Columns.VerticalAlignment = Excel.Constants.xlCenter;
            range3.Interior.Color = Color.FromArgb(192, 192, 192).ToArgb();
            range3.Columns.AutoFit();//自动调整列宽度
            //在某个区域写入数据数组
            var matrixHeight = 20;
            var matirxWidth = 20;

            xlsBook.Close();
            _xlsApp.Quit();
            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            //Console.WriteLine("花费时间是：{0}", sw.ElapsedMilliseconds * 1000.0 / Stopwatch.Frequency);
            Console.WriteLine("花费时间是：{0}", ts2.TotalMilliseconds / 1000.0f);
        }
    }
}