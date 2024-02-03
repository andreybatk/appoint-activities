using ExcelParser.Models;
using System;
using System.Collections.Generic;

namespace ExcelParser
{
    /// <summary>
    /// Достижение максимальной продуктивности насаждения для ведения лесопромышленной деятельности, путем рубок древесины по диаметру.
    /// </summary>
    class Activitie6 : IActivitie
    {
        private ExcelHelper _excelHelper;
        private Dictionary<string, int> _foundColumnsForFilling;
        private int _currentRow;

        /// <summary>
        /// Общая полнота
        /// </summary>
        private double _totalpol;
        /// <summary>
        /// Количество подроста
        /// </summary>
        private double _npdr;
        /// <summary>
        /// ОЗУ
        /// </summary>
        private int _ozu;
        /// <summary>
        /// Диаметр D1
        /// </summary>
        private int _d1;
        /// <summary>
        /// Ярус 2
        /// </summary>
        private int _jr2;

        public Activitie6(ExcelHelper excelHelper, Dictionary<string, int> foundColumnsForFilling, int currentRow)
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
            if (column == Columns.RequiredColumns[1]) // OZU
            {
                if (int.TryParse(cellText, out int result))
                {
                    _ozu = result;
                }
            }
            if (column == Columns.RequiredColumns[9]) // D1
            {
                if (int.TryParse(cellText, out int result))
                {
                    _d1 = result;
                }
            }
            if (column == Columns.RequiredColumns[3]) // NPDR
            {
                if (double.TryParse(cellText.Replace(',', '.'), out double result))
                {
                    _npdr = result;
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
            int activitieTwo = 0;
            int prvb = 0;   

            if (_totalpol > 0.5)
            {
                activitieOne = 30; // рубка по состоянию (диаметру)
                if (_d1 >= 16)
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
                else if(_d1 < 16)
                {
                    activitieOne = 581; // рубка ухода
                    switch (_totalpol)
                    {
                        case 0.6:
                            prvb = 5;
                            break;
                        case 0.7:
                            prvb = 10;
                            break;
                        case 0.8:
                            prvb = 15;
                            break;
                        case 0.9:
                            prvb = 25;
                            break;
                        case 1:
                            prvb = 30;
                            break;
                        case 1.1:
                            prvb = 30;
                            break;
                        case 1.2:
                            prvb = 35;
                            break;
                        case 1.3:
                            prvb = 40;
                            break;
                        case 1.4:
                            prvb = 40;
                            break;
                        case 1.5:
                            prvb = 45;
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (_totalpol == 0.5 || _totalpol == 0.4 || _totalpol == 0.3)
            {
                activitieOne = 55; // заключительный прием выборочных рубок
                prvb = 70; // Процент выборки 70%


                if (_jr2 == 0)
                {
                    activitieTwo = AppointActivitieTwo();
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
