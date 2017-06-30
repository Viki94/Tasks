Solutions are created with C# language.


Task 1:

Needed references for this project:
    - Microsoft Excel 16.0 Object Libarary - used for creating the Excel table and Excel sheet
    - System.Drawing - used for changing row color and text color

The project contains two classes:
    - Constants - includes all constants
    - ExcelCreator - includes project logic

Main method:
First I create an Excel application, Workbook and Worksheet

    CheckIfExcelAppIsCreated method accepts excelApp as argumnet and checks if Excel application was created

excelApp.Visible - show the created Excel file 

    CheckIfWorksheetIsCreated method accepts worksheet as argument and checks if worksheet was created

Set data variable to be an object with parameters named rows and colums

    CreateFirstRowData method accepts data as argument and create the first row of the table: Name, Age, Score and Average score

    WriteData method accepts rows, columns, worksheet and data as arguments. This method creates random data. The generated names are from hardcored list. Ages are from 20 to 80 and  scores are from 0 to 100. Define a range where the random data will be saved - 
    Range writeRange = worksheet.Range[startCell, endCell];

        UpdateOddRows method accepts worksheet as argument. This method creates an Excel formula. All odd rows text (rows with random data) are colored in green

    UpdateFirstRow method accepts worksheet as argument. This method select the first row and set its background to be blue and text to be bold.

    CreateAverageFormula method accepts worksheet as argument. This method selects the cell "E2" and write in it the average score from cell "C2" to cell "C101"

    AutoFitColumnE1 method accepts worksheet as argument. This method selects the cell "E1" and resize the column width to fit the text in cells.

    SaveFileInDirectory method accepts excelApp as argument. This method creates Excel file named scores.xlsx and saves it in the solution directory