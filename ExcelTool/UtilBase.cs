using OfficeOpenXml;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.ComponentModel;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;

namespace ExcelTool
{
    public class UtilBase
    {
        enum ImportHeaders
        {
            [IndexValue(1)] Key,
            [IndexValue(2)] TCName,
            [IndexValue(4)] Precondition,
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
            [IndexValue(5)] Precondition,
            [IndexValue(6)] Step,
            [IndexValue(7)] Result,
        }

        string _key;
        string _tcName;
        string _preCondition;
        string _objective;
        string _priority;
        string _labels;
        string _step;
        string _testData;
        string _result;
        string _folder;

        /// <summary>
        /// Provide the spreadsheet filename (without .xlsx extension) as an argument and place the file in C:\TestCases
        /// </summary>
        /// <param name="fileName"></param>
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
            int Import_Folder = ImportHeaders.Folder.GetIndex();

            int Dest_ID = DestHeaders.TestID.GetIndex();
            int Dest_TCName = DestHeaders.TCName.GetIndex();
            int Dest_Description = DestHeaders.Description.GetIndex();
            int Dest_Tags = DestHeaders.Tags.GetIndex();
            int Dest_Precondition = DestHeaders.Precondition.GetIndex();
            int Dest_Step = DestHeaders.Step.GetIndex();
            int Dest_Result = DestHeaders.Result.GetIndex();

            string worksheetName = fileName;
            string filePath = $"{Utils.GetCodeBasePath(false)}\\TestCases\\{fileName}";
            Console.WriteLine($"###FilePath: {filePath}");

            using (var destPkg = new ExcelPackage())
            {
                //Create an instance of the new workbook and worksheet
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
                Console.WriteLine(existingFile.ToString());
                using (var importPkg = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet importWS = importPkg.Workbook.Worksheets["Sheet0"];
                    int colCount = importWS.Dimension.End.Column;
                    int rowCount = importWS.Dimension.End.Row;
                    int destRow = 2;

                    for (int row = 2; row < rowCount; row++)
                    {
                        _key = importWS.Cells[row, Import_Key].Value?.ToString();
                        _tcName = importWS.Cells[row, Import_TCName].Value?.ToString();
                        _preCondition = importWS.Cells[row, Import_PreCondition].Value?.ToString();
                        _objective = importWS.Cells[row, Import_Objective].Value?.ToString();
                        _priority = importWS.Cells[row, Import_Priority].Value?.ToString();
                        _labels = importWS.Cells[row, Import_Labels].Value?.ToString();
                        _step = importWS.Cells[row, Import_Step].Value?.ToString();
                        _testData = importWS.Cells[row, Import_TestData].Value?.ToString();
                        _result = importWS.Cells[row, Import_Result].Value?.ToString();
                        _folder = importWS.Cells[row, Import_Folder].Value?.ToString();



                        if (!string.IsNullOrWhiteSpace(_key))
                        {
                            if (!string.IsNullOrWhiteSpace(_folder) && _folder.Contains("/"))
                            {
                                var folderHiearchy = Regex.Split(_folder, "/");
                                int fhCount = folderHiearchy.Length;

                                //TODO:
                                //folder(index1) - create worksheet
                                //subfolder1(index2) - if worksheet exists (else create new worksheet)
                                    //--get row count
                                    //--if merged row with subfolder name exists, append to end of row count (else create merged row)
                                //subfolder2(index3) - if (subfolder1(index2)) worksheet exists (else create new worksheet)
                                    //--get row count
                                    //--if merged row with 


                                if (fhCount > 1)
                                {
                                    //TODO: use hashmap to track sub-folder merged-cells exists
                                    //or Create new worksheet for each sub-folder entry
                                    for (int i = 2; i < fhCount; i++)
                                    {
                                        using (var mergedCell = destWS.Cells[row, 1])
                                        {
                                            mergedCell.Style.Font.Bold = true;
                                            mergedCell.Value = folderHiearchy[i];
                                            destWS.Cells[row, 1, row, 7].Merge = true;
                                            destWS.Cells[row, 1, row, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        }
                                    }
                                }
                            }

                            if (!string.IsNullOrWhiteSpace(_labels))
                            {
                                if (_labels.Contains(","))
                                {
                                    string[] multiLabel = Regex.Split(_labels, ", ");

                                    for (int i = 0; i < multiLabel.Length; i++)
                                    {
                                        destWS.Cells[destRow, Dest_Tags].Value = multiLabel[i];
                                    }
                                }
                                else
                                {
                                    destWS.Cells[destRow, Dest_Tags].Value = _labels;
                                }
                            }

                            destWS.Cells[destRow, Dest_ID].Value = _key;
                            destWS.Cells[destRow, Dest_TCName].Value = _tcName;
                            destWS.Cells[destRow, Dest_Description].Value = _objective;
                            destWS.Cells[destRow, Dest_Precondition].Value = _preCondition;

                            
                            destWS.Cells[destRow, Dest_Tags].Value = _priority;

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
