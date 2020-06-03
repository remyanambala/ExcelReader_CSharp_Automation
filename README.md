# ExcelReader_CSharp_Automation
              
A library built based on the C# ExcelDataReader that can used for reading data from excel for test automation. This loads the excel and its sheets into a static class which can be shared by multiple automation test scripts. It supports loading mutliple excel files.    
ExcelDataReader - https://github.com/ExcelDataReader/ExcelDataReader

Excel file format: .xlsx    
To read the protected excel files pass password in optional pwd parameter

# Functions:   
- ExcelReader.Load(string dataFilePath, optional string pwd)     
   Function: Checks whether excel is loaded and if not loads the excel and its sheets.     
   return: void    
- ExcelReader.ReadData(string filepath, string sheetName, int rowNumber, string columnName, optional string pwd)      
   Function: Loads the file if already not loaded. Returns the column value.    
   return: column value as string    
- ExcelReader.GetRowCount(string filepath, string sheetName, optional string pwd)       
  Function: Loads the file if already not loaded. Returns the row count.     
  return: row count as int   

Project includes unit tests for the library created using MSTest and Specflow.    
ExcelReaderMSTests => MS Test unit tests    
SpecFlowExcelReaderTests => Specflow (BDD) unit tests
