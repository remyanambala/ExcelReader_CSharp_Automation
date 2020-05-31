using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReaderMSTests
{
    [TestClass]
    public class ExcelReaderMSTests
    {
        public static string buildLoc = System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
        public static readonly String dataFilePath = Path.GetFullPath(Path.Combine( buildLoc, @"..\..\TestFiles\","TestExcel1.xlsx"));
      

        String filepath;

        [TestMethod]
        public void TestReadValueFromExcel1()
        {
            ExcelReader.OpenExcel(dataFilePath);
           Console.WriteLine(ExcelReader.ReadData(dataFilePath,"Purchase",1, "Fname"));

        }
        public void TestReadValueFromExcel2()
        {
        }
    }
}
