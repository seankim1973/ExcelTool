using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelTool
{
    public class UtilBase
    {
        public void ReadWorkbook(string filePath)
        {
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                //--Get name of the worksheet
                //string sheetname = worksheet.Name;
                //Console.WriteLine(sheetname);

                //--Get worksheet by name
                //ExcelWorksheet sheet0 = package.Workbook.Worksheets["Sheet0"];
                //Console.WriteLine($"SHEET_0 NAME: {sheet0.Name}");

                //--Get cell value [int Row, int Col]
                //string colPrecondition = worksheet.Cells[2,4].Value.ToString();
                //Console.WriteLine(colPrecondition);


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

        public void GetWorksheet<T>(string filePath, T sheetNameOrIndex)
        {
            Type argType = sheetNameOrIndex.GetType();
            FileInfo existingFile = new FileInfo(filePath);
            ExcelWorksheet worksheet = null;

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                if (argType == typeof(int))
                {
                    int index = (int)Convert.ChangeType(sheetNameOrIndex, typeof(int));
                    worksheet = package.Workbook.Worksheets[index];
                }
                else if (argType == typeof(string))
                {
                    string name = (string)Convert.ChangeType(sheetNameOrIndex, typeof(string));
                    worksheet = package.Workbook.Worksheets[name];
                }

                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                Console.WriteLine($"Column count: {colCount} \nRow count: {rowCount}");
                //for (int row = 2; row < 16; row++)
                //{
                //    var tcKey = worksheet.Cells[row, 1].Value;
                //    if (tcKey != null)
                //    {
                //        Console.WriteLine(tcKey);
                //    }

                //    for (int col = 2; col < 18; col++)
                //    {
                //        var columnData = worksheet.Cells[row, col].Value;

                //        if (columnData != null)
                //        {
                //            Console.WriteLine($"        {worksheet.Cells[1, col].Value}: {columnData}");
                //        }
                //    }
                //}
            }


            if (worksheet == null)
            {
                Assert.IsNotNull(worksheet,
                    $"Worksheet in file not found. \nFile: {filePath} \nWorksheet: {sheetNameOrIndex}");
            }

        }


        static string _key;
        static string _tcName;
        static string _preCondition;
        static string _objective;
        static string _priority;
        static string _labels;
        static string _step;
        static string _testData;
        static string _result;

        /// <summary>
        /// Provide the spreadsheet filename (without .xlsx extension) as an argument and place the file in C:\TestCases
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="worksheetName"></param>
        public void FormatWorksheet(string fileName)
        {
            int Import_Key = ImportHeaders.Key.GetIndex();
            int Import_TCName = ImportHeaders.TCName.GetIndex();
            int Import_PreCondition = ImportHeaders.Precondition.GetIndex();
            int Import_Objective = ImportHeaders.Objective.GetIndex();
            int Import_Priority = ImportHeaders.Priority.GetIndex();
            int Import_Labels = ImportHeaders.Labels.GetIndex();
            int Import_Step = ImportHeaders.Step.GetIndex();
            int Import_TestData = ImportHeaders.TestData.GetIndex();
            int Import_Result = ImportHeaders.Result.GetIndex();

            int Dest_ID = DestHeaders.TestID.GetIndex();
            int Dest_TCName = DestHeaders.TCName.GetIndex();
            int Dest_Description = DestHeaders.Description.GetIndex();
            int Dest_Tags = DestHeaders.Tags.GetIndex();
            int Dest_Precondition = DestHeaders.Precondition.GetIndex();
            int Dest_Step = DestHeaders.Step.GetIndex();
            int Dest_Result = DestHeaders.Result.GetIndex();

            string worksheetName = fileName;
            string filePath = $"C:\\TestCases\\{fileName}";

            //Create an instance of the new workbook and worksheet
            using (var destPkg = new ExcelPackage())
            {
                ExcelWorksheet destWS = destPkg.Workbook.Worksheets.Add(worksheetName);
                destWS.Cells[1, Dest_ID].Value = "Test ID";
                destWS.Cells[1, Dest_TCName].Value = "Test name";
                destWS.Cells[1, Dest_Description].Value = "Test description";
                destWS.Cells[1, Dest_Tags].Value = "Test tags";
                destWS.Cells[1, Dest_Precondition].Value = "Pre-condition";
                destWS.Cells[1, Dest_Step].Value = "Steps";
                destWS.Cells[1, Dest_Result].Value = "Result";

                //Read from worksheet to export
                FileInfo existingFile = new FileInfo($"{filePath}.xlsx");
                using (var importPkg = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet importWS = importPkg.Workbook.Worksheets[0];
                    int colCount = importWS.Dimension.End.Column;
                    int rowCount = importWS.Dimension.End.Row;
                    int destRow = 2;

                    ExcelRangeBase destStep = null;
                    ExcelRangeBase destTestData = null;
                    ExcelRangeBase destResult = null;

                    for (int row = 2; row < rowCount; row++)
                    {
                        _key = importWS.Cells[row, Import_Key].Value.ToString();

                        if (_key != null)
                        {
                            _tcName = importWS.Cells[row, Import_TCName].Value.ToString();
                            _preCondition = importWS.Cells[row, Import_PreCondition].Value.ToString();
                            _objective = importWS.Cells[row, Import_Objective].Value.ToString();
                            _priority = importWS.Cells[row, Import_Priority].Value.ToString();
                            _labels = importWS.Cells[row, Import_Labels].Value.ToString();
                            _step = importWS.Cells[row, Import_Step].Value.ToString();
                            _testData = importWS.Cells[row, Import_TestData].Value.ToString();
                            _result = importWS.Cells[row, Import_Result].Value.ToString();

                            destWS.Cells[destRow, Dest_ID].Value = _key;
                            destWS.Cells[destRow, Dest_TCName].Value = _tcName;
                            destWS.Cells[destRow, Dest_Description].Value = _objective;

                            if(_labels.Contains(","))
                            {
                                string[] multiLabel =  Regex.Split(_labels, ", ");
                                var _Value = _labels;

                                for (int i = 0; i < multiLabel.Length; i++)
                                {
                                    _Value = $"{multiLabel[i]},{Environment.NewLine}";
                                }
                                destWS.Cells[destRow, Dest_Tags].Value = $"{multiLabel[0]}";
                            }

                            var multiline = string.Concat("abcd", Environment.NewLine, "hjkll", Environment.NewLine);

                            destWS.Cells[destRow, Dest_Tags].Value = _priority;
                            destWS.Cells[destRow, Dest_Precondition].Value = _preCondition;
                            destWS.Cells[destRow, Dest_Step].Value = _step;
                            destWS.Cells[destRow, Dest_Step].Value = _testData;
                            destWS.Cells[destRow, Dest_Result].Value = _result;                           
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(_step))
                            {
                                destWS.Cells[destRow, Dest_Step].Value = _step;
                            }
                            if (string.IsNullOrWhiteSpace(_testData))
                            {
                                destWS.Cells[destRow, Dest_Step].Value = _testData;
                            }
                            if (string.IsNullOrWhiteSpace(_result))
                            {
                                destWS.Cells[destRow, Dest_Result].Value = _result;
                            }
                        }

                        destRow++;
                    }
                }

                var xlFile = Utils.GetFileInfo($"{filePath}_HipTest.xlsx");
                destPkg.SaveAs(xlFile);
                Console.WriteLine($"File saved - {xlFile}");
            }
        }

        enum ImportHeaders
        {
            [IndexValue(1)] Key,
            [IndexValue(2)] TCName,
            [IndexValue(4)] Precondition,
            [IndexValue(5)] Objective,
            [IndexValue(7)] Priority,
            [IndexValue(9)] Labels,
            [IndexValue(14)] Step,
            [IndexValue(15)] TestData,
            [IndexValue(16)] Result
        }

        enum DestHeaders
        {
            [IndexValue(1)] TestID,
            [IndexValue(2)] TCName,
            [IndexValue(3)] Description,
            [IndexValue(4)] Tags,
            [IndexValue(5)] Precondition,
            [IndexValue(6)] Step,
            [IndexValue(7)] Result,
        }
    }
}
