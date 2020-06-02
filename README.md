# ExcelReader_CSharp_Automation
Excel reader for C# automation using Excel Data Reader

A library built based on the C# ExcelDataReader to be used in test automation. This loads the excel and its sheets as static which can be used in the automation test scripts.It supports loading mutliple files.
ExcelDataReader - https://github.com/ExcelDataReader/ExcelDataReader

Functions:

- ExcelReader.Load(string dataFilePath) 
   Function: Checks whether excel is loaded and if not loads the excel and its sheets 
   return: void
- ExcelReader.ReadData(string filepath, string sheetName, int rowNumber, string columnName)
   Function: Loads the file if already not loaded. Returns the column value
   return: column value as string
- ExcelReader.GetRowCount(string filepath, string sheetName, string columnName)
  Function: Loads the file if already not loaded. Returns the row count for column name
  return: row count as int

Project includes unit tests for the library created using MSTest and Specflow
ExcelReaderMSTests => MS Test unit tests
SpecFlowExcelReaderTests => Spec flow unit tests
