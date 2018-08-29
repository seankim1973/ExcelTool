using NUnit.Framework;
using OfficeOpenXml;
using System;

namespace ExcelTool
{
    [TestFixture]
    public class UnitTest1 : UtilBase
    {
        
        readonly string filePath = $"C:\\TestCases\\BreakSheet Module.xlsx";


        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Test]
        public void TestMethod1()
        {
            UtilBase util = new UtilBase();
            util.ReadWorkbook(filePath);
        }



        [Test]
        public void ExcelData()
        {
            UtilBase util = new UtilBase();
            util.GetWorksheet(filePath, 0);
            

        }
    }
}
