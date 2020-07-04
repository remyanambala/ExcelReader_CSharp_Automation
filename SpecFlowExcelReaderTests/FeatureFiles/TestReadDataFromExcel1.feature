Feature: ExcelReaderSpecFlow
	As an user
	I want to read data from a excel sheet

@Test read data from excel
Scenario: Read a column data from excel sheet
	Given I have loaded the excel using <fileName> and <pwd>
	When I pass <sheetName> and <rowNumber> and <columnName> to read
	Then the result should be <result>

	Examples:
	| fileName        | sheetName | rowNumber | columnName | result      | pwd   |
	| TestExcel1.xlsx | Purchase  | 1         | Fname      | David        | myPwd |
	| TestExcel1.xlsx | Purchase  | 1         | Lname      | Copper Field | myPwd |
	| TestExcel1.xlsx | Purchase  | 1         | City       | New York     | myPwd |
	| TestExcel1.xlsx | Purchase  | 1         | Total      | 1100         | myPwd |
	| TestExcel2.xlsx | Vehicles  | 1         | Vehicle      | car         |  |
	| TestExcel2.xlsx | Vehicles  | 1         | Wheels      | 4        |  |
