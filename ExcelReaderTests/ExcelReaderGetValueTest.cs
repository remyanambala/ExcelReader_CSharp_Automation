using System;
using TechTalk.SpecFlow;

namespace ExcelReaderTests
{
    [Binding]
    public class ExcelReaderGetValueTest
    {
        [Given(@"I have entered the excel path")]
        public void GivenIHaveEnteredTheExcelPath()
        {
            ScenarioContext.Current.Pending();
        }
        
        [Given(@"I have entered the row and column header")]
        public void GivenIHaveEnteredTheRowAndColumnHeader()
        {
            ScenarioContext.Current.Pending();
        }
        
        [When(@"I press submit")]
        public void WhenIPressSubmit()
        {
            ScenarioContext.Current.Pending();
        }
        
        [Then(@"the result should be shown")]
        public void ThenTheResultShouldBeShown()
        {
            ScenarioContext.Current.Pending();
        }
    }
}
