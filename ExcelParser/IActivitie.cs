namespace ExcelParser
{
    internal interface IActivitie
    {
        void CheckCells(string column, string cellText);
        void FillCells();
    }
}