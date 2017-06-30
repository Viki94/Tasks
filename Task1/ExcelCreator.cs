using System;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;

namespace Task1
{
    class ExcelCreator
    {
        static void Main()
        {
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet worksheet = workbook.Worksheets[1];

            CheckIfExcelAppIsCreated(excelApp);

            excelApp.Visible = true;

            CheckIfWorksheetIsCreated(worksheet);

            var data = new object[Constants.rows, Constants.columns];
            CreateFirstRowData(data);

            WriteData(Constants.rows, Constants.columns, worksheet, data);

            UpdateFirstRow(worksheet);

            CreateAverageFormula(worksheet);
            AutoFitColumnE1(worksheet);

            SaveFileInDirectory(excelApp);
        }

        private static void CheckIfExcelAppIsCreated(Application excelApp)
        {
            if (excelApp == null)
            {
                Console.WriteLine(Constants.EXCEL_ERROR);
                return;
            }
        }

        private static void CheckIfWorksheetIsCreated(Worksheet worksheet)
        {
            if (worksheet == null)
            {
                Console.WriteLine(Constants.WORKSHEET_ERROR);
            }
        }

        private static void CreateFirstRowData(object[,] data)
        {
            data[0, 0] = Constants.NAME;
            data[0, 1] = Constants.AGE;
            data[0, 2] = Constants.SCORE;
            data[0, 4] = Constants.AVERAGE_SCORE;
        }

        private static void WriteData(int rows, int columns, Worksheet worksheet, Object[,] data)
        {
            var names = new List<string>()
            {
                "Ivan", "Valentina", "George", "Ivelina", "Peter", "Maria", "Ralica", "Teodor", "Eva", "Damian"
            };

            var random = new Random();

            for (var row = 1; row < rows; row++)
            {
                var column = 0;
                var r = random.Next(names.Count);
                data[row, column] = names[r];

                column = 1;
                r = random.Next(Constants.MIN_AGE_VALUE, Constants.MAX_AGE_VALUE);
                data[row, column] = r;

                column = 2;
                r = random.Next(Constants.MIN_SCORE_VALUE, Constants.MAX_SCORE_VALUE);
                data[row, column] = r;
            }

            Range startCell = worksheet.Cells[1, 1];
            Range endCell = worksheet.Cells[rows, columns];
            Range writeRange = worksheet.Range[startCell, endCell];

            UpdateOddRows(worksheet);

            writeRange.Value2 = data;
        }

        private static void UpdateOddRows(Worksheet worksheet)
        {
            FormatCondition format = worksheet.Rows.FormatConditions
                .Add(XlFormatConditionType.xlExpression, XlFormatConditionOperator.xlEqual, Constants.ODD_ROWS_FORMULA);

            format.Font.Color = XlRgbColor.rgbGreen;
        }

        private static void UpdateFirstRow(Worksheet worksheet)
        {
            worksheet.Rows[1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
            worksheet.Rows[1].EntireRow.Font.Bold = true;
        }

        private static void CreateAverageFormula(Worksheet worksheet)
        {
            Range range = worksheet.Range[Constants.E2];
            range.Formula = Constants.AVERAGE_FORMULA;
        }

        private static void AutoFitColumnE1(Worksheet worksheet)
        {
            Range range = worksheet.Range[Constants.E1];
            range.Columns.AutoFit();
        }

        private static void SaveFileInDirectory(Application excelApp)
        {
            string projectPath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory()));
            excelApp.Application.ActiveWorkbook.SaveAs(projectPath + Constants.FILE_PATH,
                XlSaveAsAccessMode.xlNoChange);
        }
    }
}
