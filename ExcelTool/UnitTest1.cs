using NUnit.Framework;

namespace ExcelTool
{
    [TestFixture]
    public class UnitTest1
    {
        
        readonly string filePath = $"C:\\Users\\schong\\Downloads\\atm-exporter (1).xlsx";

        [SetUp]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Test]
        public void TestMethod1()
        {
            UtilBase.ReadWorkbook(filePath);
        }

        [Test]
        public void ExcelData()
        {
            UtilBase.GetExcelData(filePath);
        }
    }
}
