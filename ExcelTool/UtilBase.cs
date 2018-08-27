using ExcelDataReader;
using ExcelNumberFormat;
using OfficeOpenXml;
using System;
using System.Data;
using System.Globalization;
using System.IO;

namespace ExcelTool
{
    public class UtilBase
    {
        public static void ReadWorkbook(string filePath)
        {
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                //string sheetname = worksheet.Name;
                //Console.WriteLine(sheetname);

                //ExcelWorksheet sheet0 = package.Workbook.Worksheets["Sheet0"];
                //Console.WriteLine($"SHEET_0 NAME: {sheet0.Name}");

                //string colPrecondition = worksheet.Cells[2,4].Value.ToString();
                //Console.WriteLine(colPrecondition);

                //List<ExcelColumn> columns = new List<ExcelColumn>();
                //List<ExcelRow> rows = new List<ExcelRow>();

                //ExcelRow keyRow = worksheet.Row(1);
                //ExcelColumn keyColumn = worksheet.Column(1);

                for (int row = 2; row < 16; row++)
                {
                    var tcKey = worksheet.Cells[row, 1].Value;
                    if (tcKey != null)
                    {
                        Console.WriteLine(tcKey);
                    }
                    
                    for (int col = 2; col < 18; col++)
                    {
                        var columnData = worksheet.Cells[row, col].Value;

                        if (columnData != null)
                        {
                            Console.WriteLine($"        {worksheet.Cells[1, col].Value}: {columnData}");
                        }
                        
                    }
                }
            }
        }

    }
}
