# ExcelReader_CSharp_Automation
              
A library built based on the C# ExcelDataReader that can be used for reading data from excel for test automation. This library loads the excel and its sheets into a static class which can be shared by multiple automation test scripts. It supports loading mutliple excel files.    
ExcelDataReader - https://github.com/ExcelDataReader/ExcelDataReader

Excel file format: .xlsx    
To read the protected excel files pass password in optional pwd parameter

# Functions:   
- ExcelReader.Load(string dataFilePath, optional string pwd)     
   Function: Checks whether excel is loaded and if not loads the excel and its sheets.     
   return: void    
- ExcelReader.ReadData(string filepath, string sheetName, int rowNumber, string columnName, optional string pwd)      
   Function: Returns the column value. filepath acts as the key for the required excel file as multiple files can be loaded at the same time.    
   return: column value as string    
- ExcelReader.GetRowCount(string filepath, string sheetName, optional string pwd)       
  Function: Loads the file if already not loaded. Returns the row count.     
  return: row count as int   

Project includes unit tests for the library created using MSTest and Specflow.    
ExcelReaderMSTests => MS Test unit tests    
SpecFlowExcelReaderTests => Specflow (BDD) unit tests

Azure devops continuous integration (CI/CD) pipeline can be found under below location. The build pipeline triggers the release flow where unit tests  - both MS test and Spec flow are run.:  
https://dev.azure.com/remyanambala/ExcelReader/    

Specflow results:   
<img src="https://github.com/remyanambala/ExcelReader_CSharp_Automation/blob/master/Resources/Specflow.png">

MSTest results:  
<img src="https://github.com/remyanambala/ExcelReader_CSharp_Automation/blob/master/Resources/MStest_results.jpg">

How to use this in selenium scripts:   
Below is a screenshot of using it in Easy Repro - MS Dynamics 365 Selenium automation framework.
<img src="https://github.com/remyanambala/ExcelReader_CSharp_Automation/blob/master/Resources/SeleniumScript.jpg">
