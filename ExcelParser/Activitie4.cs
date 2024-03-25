using ExcelParser.Models;
using System.Collections.Generic;

namespace ExcelParser
{
    /// <summary>
    /// Чересполосное пасечное ведение лесного хозяйства.
    /// </summary>
    internal class Activitie4 : IActivitie
    {
        private ExcelHelper _excelHelper;
        private Dictionary<string, int> _foundColumnsForFilling;
        private int _currentRow;
        /// <summary>
        /// Категория защитности
        /// </summary>
        private int _katl;
        /// <summary>
        /// ОЗУ
        /// </summary>
        private int _ozu;
        /// <summary>
        /// Группа возраста
        /// </summary>
        private int _grvoz;
        /// <summary>
        /// Порода 1
        /// </summary>
        private int _por1;
        /// <summary>
        /// Порода 2
        /// </summary>
        private int _por2;
        /// <summary>
        /// Коэффициент состава породы 1
        /// </summary>
        private int _ks1;
        /// <summary>
        /// Коэффициент состава породы 2
        /// </summary>
        private int _ks2;

        public Activitie4(ExcelHelper excelHelper, Dictionary<string, int> foundColumnsForFilling, int currentRow)
        {
            _excelHelper = excelHelper;
            _foundColumnsForFilling = foundColumnsForFilling;
            _currentRow = currentRow;
        }

        public void FillCells()
        {
            var appointActivitieOne = AppointActivities();

            foreach (var column in _foundColumnsForFilling)
            {
                if (column.Key == "PRVB")
                {
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item2.ToString());
                }
                if (column.Key == "MER1")
                {
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item1.ToString());
                }
            }
        }
        public void CheckCells(string column, string cellText)
        {
            if (column == Columns.RequiredColumns[0]) // KATL
            {
                if (int.TryParse(cellText, out int result))
                {
                    _katl = result;
                }
            }
            if (column == Columns.RequiredColumns[1]) // OZU
            {
                if (int.TryParse(cellText, out int result))
                {
                    _ozu = result;
                }
            }
            if (column == Columns.RequiredColumns[2]) // GR VOZ
            {
                if (int.TryParse(cellText, out int result))
                {
                    _grvoz = result;
                }
            }
            if (column == Columns.RequiredColumns[4]) // POR1
            {
                if (int.TryParse(cellText, out int result))
                {
                    _por1 = result;
                }
            }
            if (column == Columns.RequiredColumns[5]) // POR2
            {
                if (int.TryParse(cellText, out int result))
                {
                    _por2 = result;
                }
            }
            if (column == Columns.RequiredColumns[6]) // KS1
            {
                if (int.TryParse(cellText, out int result))
                {
                    _ks1 = result;
                }
            }
            if (column == Columns.RequiredColumns[7]) // KS2
            {
                if (int.TryParse(cellText, out int result))
                {
                    _ks2 = result;
                }
            }
        }
        private (int activitieOne, int prvb) AppointActivities()
        {
            if (_ozu != 0) return (0, 0);

            int activitieOne = 0;
            int prvb = 0;

            var result = (activitieOne, prvb);

            if (_grvoz == 4 || _grvoz == 5) // Выбираем спелые и перестойные насаждения 
            {
                if (_por1 == 100100 || _por1 == 100108 || _por1 == 100185 || _por1 == 100150 || _por1 == 100200
                    || _por1 == 100215 || _por1 == 100230 || _por1 == 100240 || _por1 == 100241 || _por1 == 100300
                    || _por1 == 100345 || _por1 == 100400 || _por1 == 100410 || _por1 == 100440 || _por1 == 100500)
                {
                    activitieOne = 90;
                    prvb = 50;
                }
                else
                {
                    activitieOne = 90;
                    prvb = 100;
                }

                if (_ks1 >= 8)
                {
                    if (_por1 == 100200 || _por1 == 100300)
                    {
                        if (_katl == 80)
                        {
                            activitieOne = 10;
                            prvb = 100;
                        }
                        else
                        {
                            activitieOne = 0;
                            prvb = 0;
                        }
                        result = (activitieOne, prvb);
                        return result;
                    }
                }

                if ((_por1 == 100200 || _por1 == 100300) && (_por2 == 100200 || _por2 == 100300))
                {
                    if (_ks1 + _ks2 >= 8)
                    {
                        if (_katl == 80)
                        {
                            activitieOne = 10;
                            prvb = 100;
                        }
                        else
                        {
                            activitieOne = 0;
                            prvb = 0;
                        }
                        result = (activitieOne, prvb);
                        return result;
                    }
                }
            }
            result = (activitieOne, prvb);
            return result;
        }
    }
}