using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelTool
{
    [TestFixture]
    public class UnitTest1 : UtilBase
    {      
        readonly string fileName = $"BreakSheet Module";
        string path = string.Empty;

        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Utils.OutputDir = Utils.GetCodeBasePath();
            path = $"{Utils.GetCodeBasePath(false)}\\TestCases";
        }

        [Test]
        public void TestMethod1()
        {
            string[] folder = {
                "/DIR (Daily Inspection Report) Module/DIR Document",
                "/DIR (Daily Inspection Report) Module/Inspection Deficiency Log Report",
                "/DIR (Daily Inspection Report) Module/QA DIR's Filters",
                "/DIR (Daily Inspection Report) Module",
            };

            for (int f = 0; f < folder.Length; f++)
            {
                var folderHiearchy = Regex.Split(folder[f], "/");
                int fhCount = folderHiearchy.Length;

                if (fhCount > 2)
                {
                    for (int i = 2; i < fhCount; i++)
                    {
                        Console.WriteLine(folderHiearchy[i]);
                    }
                }
            }

            try
            {
                string worksheetName = "My (Worksheet)";
                
                Console.WriteLine($"PATH: {path}");
                Console.WriteLine($"OUTPUT_DIR: {Utils.OutputDir.ToString()}");
                Console.WriteLine($"BASE: {Utils.GetCodeBasePath(false).ToString()}");

                using (var destPkg = new ExcelPackage())
                {                    
                    ExcelWorksheet destWS = destPkg.Workbook.Worksheets.Add(worksheetName);
                    Console.WriteLine("Created an instance of the new workbook and worksheet");



                    string[] headerVals = {
                        "Test ID",
                        "Test name",
                        "Test description",
                        "Test tags",
                        "Pre-condition",
                        "Steps",
                        "Result"
                    };

                    ExcelStyle headerStyle = GetFormatStyle();

                    for (int col = 1; col < 8; col++)
                    {
                        int hValIndex = col - 1;
                        Console.WriteLine(headerVals[hValIndex]);
                        using (var headerItem = destWS.Cells[1, col])
                        {
                            
                            headerItem.Value = headerVals[hValIndex];
                            headerItem.Style.Font.Size = headerStyle.Font.Size;
                            headerItem.Style.Font.Name = headerStyle.Font.Name;
                            headerItem.Style.Hidden = headerStyle.Hidden;
                            headerItem.Style.XfId = headerStyle.XfId;
                        }
                    }


                    using (var mergedCell = destWS.Cells[2, 1])
                    {
                        mergedCell.Style.Font.Bold = true;
                        mergedCell.Value = worksheetName;
                        destWS.Cells[2, 1, 2, 7].Merge = true;
                        destWS.Cells[2, 1, 2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    var xlFile = Utils.GetFileInfo($"{worksheetName}_Hiptest.xlsx");
                    Console.WriteLine($"File saved - {xlFile}");
                    destPkg.SaveAs(xlFile);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }

        }

        [Test]
        public ExcelStyle GetFormatStyle()
        {
            string worksheetName = "excel_import_template";
            ExcelStyle style = null;

            //Read from worksheet to export
            FileInfo existingFile = new FileInfo($"{path}\\{worksheetName}.xlsx");
            using (var importPkg = new ExcelPackage(existingFile))
            {
                ExcelWorksheet importWS = importPkg.Workbook.Worksheets[0];

                using (var headerRow = importWS.Cells[1, 1])
                {
                    style = headerRow.Style;
                }
            }

            return style;
        }

        [Test]
        public void ExcelData()
        {
            //string startupPath = Directory.GetCurrentDirectory();
            //Console.WriteLine($"#1: {startupPath}");
            //startupPath = Environment.CurrentDirectory;
            //Console.WriteLine($"#2: {startupPath}");

            UtilBase util = new UtilBase();
            util.FormatWorksheet(fileName);
        }
    }
}
