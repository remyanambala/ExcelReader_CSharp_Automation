Feature: ExcelReader
	As an user
	I want to read data from a excel sheet

@mytag
Scenario: Read a column data from excel sheet
	Given I have entered the excel path
	And I have entered the row and column header
	When I press submit
	Then the result should be shown
