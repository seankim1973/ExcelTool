using NUnit.Framework;

namespace ExcelTool
{
    [TestFixture]
    public class UnitTest1
    {
        [Test]
        public void TestMethod1()
        {
            string filePath = $"C:\\Users\\schong\\Downloads\\atm-exporter (1).xlsx";
            UtilBase.ReadWorkbook(filePath);
        }
    }
}
