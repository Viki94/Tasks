namespace Task1
{
    public class Constants
    {
        public const int rows = 100001;
        public const int columns = 5;

        public const int MIN_AGE_VALUE = 20;
        public const int MAX_AGE_VALUE = 81;

        public const int MIN_SCORE_VALUE = 0;
        public const int MAX_SCORE_VALUE = 101;

        public const string EXCEL_ERROR = "EXCEL could not be started. Check that your office installation and project references are correct.";
        public const string WORKSHEET_ERROR = "Worksheet could not be created. Check that your office installation and project references are correct.";

        public const string NAME = "Name";
        public const string AGE = "Age";
        public const string SCORE = "Score";
        public const string AVERAGE_SCORE = "Average Score";

        public const string ODD_ROWS_FORMULA = "=IF(ROW() <> 1; MOD(ROW();2) = 1)";

        public const string E2 = "E2";
        public const string AVERAGE_FORMULA = "=AVERAGE(C2:C101)";

        public const string E1 = "E1";

        public const string FILE_PATH = "@..\\..\\scores.xlsx";
    }
}
