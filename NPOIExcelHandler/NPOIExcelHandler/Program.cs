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

            if (File.Exists(tmpPath))
            {
                File.Delete(tmpPath);
            }
            
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
            
            // tmpPath = @"D:\C_Spaces\ToolsLibrary\GenerateArea\Excel2007.xls";
            //
            // var workbook2007 = new XSSFWorkbook();
            // workbook2007.CreateSheet("Sheet1");
            // workbook2007.CreateSheet("Sheet2");
            // workbook2007.CreateSheet("Sheet3");
            // var fs2007 = new FileStream(tmpPath , FileMode.Create);
            // workbook2007.Write(fs2007);
            // fs2007.Close();
            // workbook2007.Close();
            
            var workbook2003 = new HSSFWorkbook();
            workbook2003.CreateSheet("Sheet1");
            var sheetOne = (HSSFSheet)workbook2003.GetSheet("Sheet1");
            for (int i = 0; i < 10; i++)
            {
                sheetOne.CreateRow(i);
            }

            var sheetRow = (HSSFRow) sheetOne.GetRow(0);
            var sheetCell = new HSSFCell[10];
            for (int i = 0; i < 10; i++)
            {
                sheetCell[i] = (HSSFCell) sheetRow.CreateCell(i);
            }
            
            sheetCell[0].SetCellValue(true);
            sheetCell[1].SetCellValue(0.111111);
            sheetCell[2].SetCellValue("Excel2003");
            sheetCell[3].SetCellValue("123456789132456798");

            for (int i = 4; i < 10; i++)
            {
                sheetCell[i].SetCellValue(i);
            }
            
            var fs = new FileStream(tmpPath, FileMode.Create);
            workbook2003.Write(fs);
            fs.Close();
            workbook2003.Close();

        }
    }
}