using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Remya.ExcelReader;

namespace ExcelReaderMSTests
{
    [TestClass]
    public class ExcelReaderMSTests
    {
        public string buildLoc = System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
        public String dataFilePath;


        [TestMethod]
        public void TestReadValueFromExcel1()
        {
            dataFilePath = Path.GetFullPath(Path.Combine(buildLoc, @"..\..\..\TestFiles\", "TestExcel1.xlsx"));
            ExcelReader.Load(dataFilePath);
            string fname = ExcelReader.ReadData(dataFilePath, "Purchase", 1, "Fname");
            string lname = ExcelReader.ReadData(dataFilePath, "Purchase", 1, "Lname");
            string city = ExcelReader.ReadData(dataFilePath, "Purchase", 1, "City");
            string total = ExcelReader.ReadData(dataFilePath, "Purchase", 1, "Total");
            Console.WriteLine($"{fname} {lname} from {city} has spend {total}");
            Assert.AreEqual("David", fname, $"First name is not correct: Expected: David, Actual: {fname}");
            Assert.AreEqual("Copper Field", lname, $"Last name is not correct: Expected: Copper Field, Actual: {lname}");
            Assert.AreEqual("New York", city, $"City name is not correct: Expected: New York, Actual: {city}");
            Assert.AreEqual("1100", total, $"Total is not correct: Expected: 1100, Actual: {total}");

        }
        [TestMethod]
        public void TestLoadDuringCountAndLoop()
        {
            dataFilePath = Path.GetFullPath(Path.Combine(buildLoc, @"..\..\..\TestFiles\", "TestExcel2.xlsx"));
            int rowCount = ExcelReader.GetRowCount(dataFilePath, "Vehicles", "Vehicle");
            for (int row = 0; row <= rowCount-1; row++) 
             { 
                string vName = ExcelReader.ReadData(dataFilePath, "Vehicles", row, "Vehicle");
                string wheels = ExcelReader.ReadData(dataFilePath, "Vehicles", row, "wheels");
                Console.WriteLine($"{vName} have {wheels} wheels");
            }

        }
    }
}
