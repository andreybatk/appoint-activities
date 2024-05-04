using System;
using AppointActivities.DB;
using AppointActivities.WPF.Infrastructure;

namespace AppointActivities.WPF.Models
{
    /// <summary>
    /// Третий сценарий. Назначения мероприятий - максимально лесокультурный эффект
    /// </summary>
    internal class Activitiy3 : IActivitie
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

        public Activitiy3(OS_INFO currentRow)
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
            if (_currentRow.Gr_voz == 4 || _currentRow.Gr_voz == 5)
            {
                if (_currentRow.KATL == 80)
                {
                    if (_currentRow.OZU != 0) return;

                    if ((_currentRow.TLU >= 1 && _currentRow.TLU <= 4) || (_currentRow.TLU >= 7 && _currentRow.TLU <= 10) || (_currentRow.TLU >= 13 && _currentRow.TLU <= 16))
                    {
                        _mer1 = 10; // сплошные рубки
                        _prvb = 100;
                        _mer2 = AppointMer2();
                    }
                    else
                    {
                        _mer1 = 80; // выборочные рубки

                        CalculatePol();
                        if (_totalPol == 0.5 || _totalPol == 0.4 || _totalPol == 0.3)
                        {
                            _mer1 = 55; // заключительный прием выборочных рубок
                            _prvb = 70; // Процент выборки 70%

                            if (_currentRow.JR2 == 0)
                            {
                                _mer2 = AppointMer2();
                            }
                        }
                        else
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
                    }
                }

                if (_currentRow.KATL != 80) //защитные
                {
                    if (_currentRow.KS1 >= 8)
                    {
                        if (_currentRow.POR1 == 100200 || _currentRow.POR1 == 100300) //темная хвоя ель пихта
                        {
                            _mer1 = 0;
                            _mer2 = 0;
                            _prvb = 0;
                        }
                    }

                    if ((_currentRow.POR1 == 100200 || _currentRow.POR1 == 100300) && (_currentRow.POR2 == 100200 || _currentRow.POR2 == 100300))
                    {
                        if (_currentRow.KS1 + _currentRow.KS1 >= 8)
                        {
                            _mer1 = 0;
                            _mer2 = 0;
                            _prvb = 0;
                        }
                    }
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