using System;
using TechTalk.SpecFlow;
using Remya.ExcelReader;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SpecFlowExcelReaderTests
{
    [Binding]
    public class ExcelReaderTestSteps
    {
        string result;
        string buildLoc = System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
        string dataFilePath;
       [Given(@"I have loaded the excel using (.*) and (.*)")]
        public void GivenIHaveLoadedTheExcelUsingAnd(string fileName, string pwd = null)
        {
            //ScenarioContext.Current.Pending();
            
            dataFilePath = Path.GetFullPath(Path.Combine(buildLoc, @"..\..\..\TestFiles\", fileName));
            ExcelReader.Load(dataFilePath, pwd);
        }

        [When(@"I pass (.*) and (.*) and (.*) to read")]
        public void WhenIPassAndAndToRead(string sheetName, int rowNumber, string columnName)
        {
            result =ExcelReader.ReadData(dataFilePath, sheetName, Convert.ToInt32(rowNumber), columnName);
        }
        
        [Then(@"the result should be (.*)")]
        public void ThenTheResultShouldBe(string expResult)
        {
           // this.result.Equals(expResult);
            Assert.AreEqual(expResult, this.result, $"values don't match actual = {result}, expected = {expResult}");
        }
    }
}
