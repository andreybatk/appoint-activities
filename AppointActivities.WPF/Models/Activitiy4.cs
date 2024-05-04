using AppointActivities.DB;
using AppointActivities.WPF.Infrastructure;

namespace AppointActivities.WPF.Models
{
    /// <summary>
    /// Четвертый сценарий. Чересполосное пасечное ведение лесного хозяйства.
    /// </summary>
    internal class Activitiy4 : IActivitie
    {
        /// <summary>
        /// Мероприятие 1
        /// </summary>
        private int _mer1 = 0;
        /// <summary>
        /// Процент выборки
        /// </summary>
        private int _prvb = 0;
        /// <summary>
        /// Текущий элемент из БД
        /// </summary>
        //private MyTable _currentRow;
        private OS_INFO _currentRow;

        public Activitiy4(OS_INFO currentRow)
        {
            _currentRow = currentRow;
        }

        public void AppointActivitie()
        {
            _currentRow.PRVB = _prvb;
            _currentRow.MER1 = _mer1;
        }
        public void CalculateActivitie()
        {
            if (_currentRow.OZU != 0) return;

            if (_currentRow.Gr_voz == 4 || _currentRow.Gr_voz == 5) // Выбираем спелые и перестойные насаждения 
            {
                if (_currentRow.POR1 == 100100 || _currentRow.POR1 == 100108 || _currentRow.POR1 == 100185 || _currentRow.POR1 == 100150 || _currentRow.POR1 == 100200
                    || _currentRow.POR1 == 100215 || _currentRow.POR1 == 100230 || _currentRow.POR1 == 100240 || _currentRow.POR1 == 100241 || _currentRow.POR1 == 100300
                    || _currentRow.POR1 == 100345 || _currentRow.POR1 == 100400 || _currentRow.POR1 == 100410 || _currentRow.POR1 == 100440 || _currentRow.POR1 == 100500)
                {
                    _mer1 = 90;
                    _prvb = 50;
                }
                else
                {
                    _mer1 = 90;
                    _prvb = 100;
                }

                if (_currentRow.KS1 >= 8)
                {
                    if (_currentRow.POR1 == 100200 || _currentRow.POR1 == 100300)
                    {
                        if (_currentRow.KATL == 80)
                        {
                            _mer1 = 10;
                            _prvb = 100;
                        }
                        else
                        {
                            _mer1 = 0;
                            _prvb = 0;
                        }
                        return;
                    }
                }

                if ((_currentRow.POR1 == 100200 || _currentRow.POR1 == 100300) && (_currentRow.POR2 == 100200 || _currentRow.POR2 == 100300))
                {
                    if (_currentRow.KS1 + _currentRow.KS2 >= 8)
                    {
                        if (_currentRow.KATL == 80)
                        {
                            _mer1 = 10;
                            _prvb = 100;
                        }
                        else
                        {
                            _mer1 = 0;
                            _prvb = 0;
                        }
                        return;
                    }
                }
            }
        }
    }
}