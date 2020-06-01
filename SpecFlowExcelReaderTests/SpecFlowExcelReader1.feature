Feature: ExcelReader
	As an user
	I want to read data from a excel sheet

@mytag
Scenario: Read a column data from excel sheet
	Given I have loaded the excel using <fileName>
	When I pass <sheetName> and <rowNumber> and <columnName> to read
	Then the result should be <result>

	Examples:
	| fileName        | sheetName | rowNumber | columnName | result       |
	| TestExcel1.xlsx | Purchase  | 0         | Fname      | David        |
	| TestExcel1.xlsx | Purchase  | 0         | Lname      | Copper Field |
	| TestExcel1.xlsx | Purchase  | 0         | City       | New York     |
	| TestExcel1.xlsx | Purchase  | 0         | Total      | 1100         |
