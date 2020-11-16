using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOIExcelHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            
            var tmpPath = @"D:\C_Spaces\ToolsLibrary\GenerateArea\Excel2003.xls";
            
            // if (!File.Exists(tmpPath))
            // {
            //     var workbook2003 = new HSSFWorkbook();
            //     workbook2003.CreateSheet("Sheet1");
            //     workbook2003.CreateSheet("Sheet2");
            //     workbook2003.CreateSheet("Sheet3");
            //     var file2003 = new FileStream(tmpPath , FileMode.Create);
            //     workbook2003.Write(file2003);
            //     file2003.Close();
            //     workbook2003.Close();
            // }
            
            tmpPath = @"D:\C_Spaces\ToolsLibrary\GenerateArea\Excel2007.xls";
            
            var workbook2007 = new XSSFWorkbook();
            workbook2007.CreateSheet("Sheet1");
            workbook2007.CreateSheet("Sheet2");
            workbook2007.CreateSheet("Sheet3");
            var fs2007 = new FileStream(tmpPath , FileMode.Create);
            workbook2007.Write(fs2007);
            fs2007.Close();
            workbook2007.Close();
            
        }
    }
}