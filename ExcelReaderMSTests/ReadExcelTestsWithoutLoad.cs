using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Remya.ExcelReader;

namespace ExcelReaderMSTests
{
    [TestClass]
    public class ExcelReaderMSTestsWithoutLoad
    {
        public static string buildLoc = System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
        public static readonly String dataFilePath = Path.GetFullPath(Path.Combine( buildLoc, @"..\..\..\TestFiles\", "TestExcel1.xlsx"));
      

        [TestMethod]
        public void TestReadValueFromExcel1Withoutload()
        {
            //ExcelReader.OpenExcel(dataFilePath);
            string fname = ExcelReader.ReadData(dataFilePath, "Purchase", 2, "Fname");
            string lname = ExcelReader.ReadData(dataFilePath, "Purchase", 2, "Lname");
            string city = ExcelReader.ReadData(dataFilePath, "Purchase", 2, "City");
            string total = ExcelReader.ReadData(dataFilePath, "Purchase", 2, "Total");
            Console.WriteLine($"{fname} {lname} from {city} has spend {total}");
         }
       
    }
}
