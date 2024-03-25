using System;
using WpfApplication.DB;
using WpfApplication.Infrastructure;

namespace WpfApplication.Models
{
    /// <summary>
    /// Пятый сценарий. Достижение максимальной продуктивности насаждения для ведения лесопромышленной деятельности, путем рубок древесины по диаметру.
    /// </summary>
    internal class Activitiy5 : IActivitie
    {
        /// <summary>
        /// Мероприятие 1
        /// </summary>
        private int _mer1 = 0;
        /// <summary>
        /// Мероприятие 2
        /// </summary>
        private int _mer2 = 0;
        /// <summary>
        /// Процент выборки
        /// </summary>
        private int _prvb = 0;
        /// <summary>
        /// Общая полнота
        /// </summary>
        private double _totalPol;
        /// <summary>
        /// Текущий элемент из БД
        /// </summary>
        //private MyTable _currentRow;
        private OS_INFO _currentRow;

        public Activitiy5(OS_INFO currentRow)
        {
            _currentRow = currentRow;
        }

        public void AppointActivitie()
        {
            _currentRow.PRVB = _prvb;
            _currentRow.MER1 = _mer1;
            _currentRow.MER2 = _mer2;
        }
        public void CalculateActivitie()
        {
            if (_currentRow.OZU != 0) return;

            CalculatePol();
            if (_totalPol > 0.5)
            {
                _mer1 = 30; // рубка по состоянию (диаметру)
                if (_currentRow.D1 >= 24)
                {
                    switch (_totalPol)
                    {
                        case 0.6:
                            _prvb = 15;
                            break;
                        case 0.7:
                            _prvb = 25;
                            break;
                        case 0.8:
                            _prvb = 35;
                            break;
                        case 0.9:
                            _prvb = 45;
                            break;
                        case 1:
                            _prvb = 50;
                            break;
                        case 1.1:
                            _prvb = 50;
                            break;
                        case 1.2:
                            _prvb = 55;
                            break;
                        case 1.3:
                            _prvb = 60;
                            break;
                        case 1.4:
                            _prvb = 60;
                            break;
                        case 1.5:
                            _prvb = 65;
                            break;
                        default:
                            break;
                    }
                }
                else if (_currentRow.D1 < 24)
                {
                    switch (_totalPol)
                    {
                        case 0.6:
                            _prvb = 5;
                            break;
                        case 0.7:
                            _prvb = 10;
                            break;
                        case 0.8:
                            _prvb = 15;
                            break;
                        case 0.9:
                            _prvb = 25;
                            break;
                        case 1:
                            _prvb = 30;
                            break;
                        case 1.1:
                            _prvb = 30;
                            break;
                        case 1.2:
                            _prvb = 35;
                            break;
                        case 1.3:
                            _prvb = 40;
                            break;
                        case 1.4:
                            _prvb = 40;
                            break;
                        case 1.5:
                            _prvb = 45;
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (_totalPol == 0.5 || _totalPol == 0.4 || _totalPol == 0.3)
            {
                _mer1 = 55; // заключительный прием выборочных рубок
                _prvb = 70; // Процент выборки 70%

                if (_currentRow.JR2 == 0)
                {
                    _mer2 = AppointMer2();
                }
            }
        }
        private int AppointMer2()
        {
            if (_currentRow.NPDR >= 0 && _currentRow.NPDR < 1)
            {
                return 500;
            }
            else if (_currentRow.NPDR >= 1 && _currentRow.NPDR < 1.5m)
            {
                return 640;
            }
            else if (_currentRow.NPDR >= 1.5m && _currentRow.NPDR < 2.5m)
            {
                return 690;
            }
            else if (_currentRow.NPDR >= 2.5m)
            {
                return 660;
            }

            return 0;
        }
        private void CalculatePol()
        {
            var pol = _currentRow.POL1 + _currentRow.POL2 + _currentRow.POL3
                + _currentRow.POL4 + _currentRow.POL5 + _currentRow.POL6
                + _currentRow.POL7 + _currentRow.POL8 + _currentRow.POL9
                + _currentRow.POL10;
            _totalPol = Convert.ToDouble(pol);
        }
    }
}