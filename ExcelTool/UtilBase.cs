using OfficeOpenXml;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.ComponentModel;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;
using System.Collections.Generic;

namespace ExcelTool
{
    public class UtilBase
    {
        enum ImportHeaders
        {
            [IndexValue(1)] Key,
            [IndexValue(2)] TCName,
            [IndexValue(4)] Preconditions,
            [IndexValue(5)] Objective,
            [IndexValue(6)] Folder,
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
            [IndexValue(5)] Preconditions,
            [IndexValue(6)] Step,
            [IndexValue(7)] Result,
        }

        string _key;
        string _tcName;
        string _preConditions;
        string _objective;
        string _priority;
        string _labels;
        string _step;
        string _testData;
        string _result;
        string _folder;

        ExcelWorksheet destWS;
        IList<string> destWsNamesList;
        static int destRow = 1;

        /// <summary>
        /// Provide the spreadsheet filename (without .xlsx extension) as an argument and place the file in C:\TestCases
        /// </summary>
        /// <param name="fileName"></param>
        public void FormatWorksheet(string fileName)
        {
            int Import_Key = ImportHeaders.Key.GetIndex();
            int Import_TCName = ImportHeaders.TCName.GetIndex();
            int Import_PreConditions = ImportHeaders.Preconditions.GetIndex();
            int Import_Objective = ImportHeaders.Objective.GetIndex();
            int Import_Priority = ImportHeaders.Priority.GetIndex();
            int Import_Labels = ImportHeaders.Labels.GetIndex();
            int Import_Step = ImportHeaders.Step.GetIndex();
            int Import_TestData = ImportHeaders.TestData.GetIndex();
            int Import_Result = ImportHeaders.Result.GetIndex();
            int Import_Folder = ImportHeaders.Folder.GetIndex();

            int Dest_ID = DestHeaders.TestID.GetIndex();
            int Dest_TCName = DestHeaders.TCName.GetIndex();
            int Dest_Description = DestHeaders.Description.GetIndex();
            int Dest_Tags = DestHeaders.Tags.GetIndex();
            int Dest_Preconditions = DestHeaders.Preconditions.GetIndex();
            int Dest_Step = DestHeaders.Step.GetIndex();
            int Dest_Result = DestHeaders.Result.GetIndex();

            string worksheetName = fileName;
            string filePath = $"{Utils.GetCodeBasePath(false)}\\TestCases\\{fileName}";
            Console.WriteLine($"###FilePath: {filePath}");

            using (var destPkg = new ExcelPackage())
            {
                //Read from worksheet to export
                FileInfo existingFile = new FileInfo($"{filePath}.xlsx");
                Console.WriteLine(existingFile.ToString());
                using (var importPkg = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet importWS = importPkg.Workbook.Worksheets["Sheet0"];
                    int colCount = importWS.Dimension.End.Column;
                    int rowCount = importWS.Dimension.End.Row;

                    for (int row = 2; row < rowCount; row++)
                    {
                        _key = importWS.Cells[row, Import_Key].Value?.ToString();
                        _tcName = importWS.Cells[row, Import_TCName].Value?.ToString();
                        _preConditions = importWS.Cells[row, Import_PreConditions].Value?.ToString();
                        _objective = importWS.Cells[row, Import_Objective].Value?.ToString();
                        _priority = importWS.Cells[row, Import_Priority].Value?.ToString();
                        _labels = importWS.Cells[row, Import_Labels].Value?.ToString();
                        _step = importWS.Cells[row, Import_Step].Value?.ToString();
                        _testData = importWS.Cells[row, Import_TestData].Value?.ToString();
                        _result = importWS.Cells[row, Import_Result].Value?.ToString();
                        _folder = importWS.Cells[row, Import_Folder].Value?.ToString();

                        if (!string.IsNullOrWhiteSpace(_key))
                        {
                            var folders = Regex.Split(_folder, "/");
                            int fhCount = (folders.Length > 2) ? 3 : folders.Length;

                            var folderHiearchy = _folder.Split(new[] { '/' }, fhCount);
                            string rootFolder = folderHiearchy[1];
                            Console.WriteLine($"ROOT FOLDER: {rootFolder}");
                            var destWorksheets = destPkg.Workbook.Worksheets;
                            destWsNamesList = new List<string>();

                            for (int i = 0; i < destWorksheets.Count; i++)
                            {
                                destWsNamesList.Add(destWorksheets[i].Name);                                
                            }

                            if (rootFolder.Length > 31)
                            {
                                rootFolder = rootFolder.Substring(0, 31);
                                Console.WriteLine($"TRUNCATED ROOT FOLDER: {rootFolder}");
                            }
                         
                            bool rootWorksheetExists = (destWsNamesList.Contains(rootFolder)) ? true : false;
                            string rootwsMsg;
                            if (rootWorksheetExists)
                            {
                                rootwsMsg = "WORKSHEET EXISTS:";
                                destWS = destPkg.Workbook.Worksheets[rootFolder];

                                if (fhCount > 2)
                                {
                                    using (var mergedCell = destWS.Cells[destRow, 1])
                                    {
                                        mergedCell.Value = folderHiearchy[2];
                                        destWS.Cells[destRow, 1, destRow, 7].Merge = true;
                                        destWS.Cells[destRow, 1, destRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        destRow++;
                                    }
                                }
                            }
                            else
                            {
                                destRow = 1;
                                rootwsMsg = "WORKSHEET NAME ADDED:";
                                destWS = destPkg.Workbook.Worksheets.Add(rootFolder);
                                destWS.Cells[destRow, Dest_ID].Value = "Test ID";
                                destWS.Cells[destRow, Dest_TCName].Value = "Test name";
                                destWS.Cells[destRow, Dest_Description].Value = "Test description";
                                destWS.Cells[destRow, Dest_Tags].Value = "Test tags";
                                destWS.Cells[destRow, Dest_Preconditions].Value = "Pre-conditions";
                                destWS.Cells[destRow, Dest_Step].Value = "Steps";
                                destWS.Cells[destRow, Dest_Result].Value = "Result";
                                destRow++;

                                if (fhCount > 2)
                                {
                                    using (var mergedCell = destWS.Cells[destRow, 1])
                                    {
                                        mergedCell.Value = folderHiearchy[2];
                                        destWS.Cells[destRow, 1, destRow, 7].Merge = true;
                                        destWS.Cells[destRow, 1, destRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        destRow++;
                                    }
                                }

                            }
                            Console.WriteLine($"{rootwsMsg} {rootFolder}");



                            destWS.Cells[destRow, Dest_ID].Value = _key;
                            destWS.Cells[destRow, Dest_TCName].Value = _tcName;
                            destWS.Cells[destRow, Dest_Description].Value = _objective;
                            destWS.Cells[destRow, Dest_Preconditions].Value = _preConditions;                           
                            destWS.Cells[destRow, Dest_Tags].Value = _priority;

                            if (!string.IsNullOrWhiteSpace(_labels))
                            {
                                if (_labels.Contains(","))
                                {
                                    string[] multiLabel = Regex.Split(_labels, ", ");

                                    StringBuilder sb = new StringBuilder();

                                    for (int i = 0; i < multiLabel.Length; i++)
                                    {
                                        sb.Append($"{multiLabel[i]}{Environment.NewLine}");
                                    }

                                    destWS.Cells[destRow, Dest_Tags].Value = sb.ToString();
                                }
                                else
                                {
                                    destWS.Cells[destRow, Dest_Tags].Value = _labels;
                                }
                            }

                            if (string.IsNullOrWhiteSpace(_step))
                            {
                                if (!string.IsNullOrWhiteSpace(_testData))
                                {
                                    destWS.Cells[destRow, Dest_Step].Value = _testData;
                                }
                            }
                            else 
                            {
                                destWS.Cells[destRow, Dest_Step].Value = (!string.IsNullOrWhiteSpace(_testData)) ?
                                    $"{_step}, {Environment.NewLine}, {_testData}" : destWS.Cells[destRow, Dest_Step].Value = _step;                            
                            }

                            destWS.Cells[destRow, Dest_Result].Value = _result;                           
                        }
                        else
                        {
                            _step = importWS.Cells[row, Import_Step].Value?.ToString();
                            _testData = importWS.Cells[row, Import_TestData].Value?.ToString();
                            _result = importWS.Cells[row, Import_Result].Value?.ToString();

                            if (!string.IsNullOrWhiteSpace(_step))
                            {
                                destWS.Cells[destRow, Dest_Step].Value = _step;
                            }
                            if (!string.IsNullOrWhiteSpace(_testData))
                            {
                                destWS.Cells[destRow, Dest_Step].Value = _testData;
                            }
                            if (!string.IsNullOrWhiteSpace(_result))
                            {
                                destWS.Cells[destRow, Dest_Result].Value = _result;
                            }
                        }

                        destRow++;
                    }
                }

                var xlFile = Utils.GetFileInfo($"{worksheetName}_Hiptest.xlsx");              
                Console.WriteLine($"File saved - {xlFile}");
                destPkg.SaveAs(xlFile);
            }
        }


    }
}
