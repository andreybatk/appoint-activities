using System.Collections.Generic;
using AppointActivities.Models;

namespace AppointActivities.CLI
{
    /// <summary>
    /// Первый сценарий назначения мероприятий - классическое хозяйство
    /// </summary>
    internal class Activitie : IActivitie
    {
        private ExcelHelper _excelHelper;
        private Dictionary<string, int> _foundColumnsForFilling;
        private int _currentRow;
        /// <summary>
        /// Общая полнота
        /// </summary>
        private double _totalpol;
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
        /// Количество подроста
        /// </summary>
        private double _npdr;
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
        /// <summary>
        /// Ярус 2
        /// </summary>
        private int _jr2;

        public Activitie(ExcelHelper excelHelper, Dictionary<string, int> foundColumnsForFilling, int currentRow)
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
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item3.ToString());
                }
                if (column.Key == "MER1")
                {
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item1.ToString());
                }
                if (column.Key == "MER2")
                {
                    var appointActivitieTwo = AppointActivitieTwo();
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item2.ToString());
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
            if (column == Columns.RequiredColumns[3]) // NPDR
            {
                if (double.TryParse(cellText.Replace(',', '.'), out double result))
                {
                    _npdr = result;
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
            if (column == Columns.RequiredJrColumns[1]) // JR2
            {
                if (int.TryParse(cellText, out int result))
                {
                    _jr2 = result;
                }
            }
            if (Columns.RequiredPolColumns.Contains(column)) //Общая полнота
            {
                if (double.TryParse(cellText, out double result))
                {
                    _totalpol += result;
                }
            }
        }
        private (int activitieOne, int activitieTwo, int prvb) AppointActivities()
        {
            if (_ozu != 0) return (0, 0, 0);

            int activitieOne = 0;
            int prvb = 0;
            int activitieTwo = 0;

            if (_grvoz == 4 || _grvoz == 5) // Выбираем спелые и перестойные насаждения 
            {

                if (_katl == 80) //эксплаутационные
                {
                    activitieOne = 10; // сплошные рубки
                    prvb = 100;
                    activitieTwo = AppointActivitieTwo();
                }
                // все кроме 80
                else
                {
                    activitieOne = 80; // выборочные рубки

                    if (_totalpol == 0.5 || _totalpol == 0.4 || _totalpol == 0.3)
                    {
                        activitieOne = 55; // заключительный прием выборочных рубок
                        prvb = 70; // Процент выборки 70%

                        if (_jr2 == 0)
                        {
                            activitieTwo = AppointActivitieTwo();
                        }
                    }
                    else
                    {
                        switch (_totalpol)
                        {
                            case 0.6:
                                prvb = 15;
                                break;
                            case 0.7:
                                prvb = 25;
                                break;
                            case 0.8:
                                prvb = 35;
                                break;
                            case 0.9:
                                prvb = 45;
                                break;
                            case 1:
                                prvb = 50;
                                break;
                            case 1.1:
                                prvb = 50;
                                break;
                            case 1.2:
                                prvb = 55;
                                break;
                            case 1.3:
                                prvb = 60;
                                break;
                            case 1.4:
                                prvb = 60;
                                break;
                            case 1.5:
                                prvb = 65;
                                break;
                            default:
                                break;
                        }
                    }
                }

                if (_katl != 80) //защитные
                {
                    if (_ks1 >= 8)
                    {
                        if (_por1 == 100200 || _por1 == 100300) //темная хвоя ель пихта
                        {
                            return (0, 0, 0);
                        }
                    }

                    if ((_por1 == 100200 || _por1 == 100300) && (_por2 == 100200 || _por2 == 100300))
                    {
                        if (_ks1 + _ks2 >= 8)
                        {
                            return (0, 0, 0);
                        }
                    }
                }
            }
            var result = (activitieOne, activitieTwo, prvb);
            return result;
        }
        private int AppointActivitieTwo()
        {
            if (_npdr >= 0 && _npdr < 1)
            {
                return 500;
            }
            else if (_npdr >= 1 && _npdr < 1.5)
            {
                return 640;
            }
            else if (_npdr >= 1.5 && _npdr < 2.5)
            {
                return 690;
            }
            else if (_npdr >= 2.5)
            {
                return 660;
            }

            return 0;
        }
    }
}