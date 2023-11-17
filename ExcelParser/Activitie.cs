using System.Collections.Generic;
using ExcelParser.Models;

namespace ExcelParser
{
    /// <summary>
    /// Назначения мероприятий: Мер1, Мер2, PRVB
    /// </summary>
    class Activitie
    {
        private ExcelHelper _excelHelper;
        private Dictionary<string, int> _foundColumnsForSettings;

        private int _currentRow;
        private bool _isAppointActivitieOne = false;
        private bool _isAppointActivitieTwo = false;

        /// <summary>
        /// Категория земель
        /// </summary>
        private int _katl { get; set; }
        /// <summary>
        /// ОЗУ
        /// </summary>
        private int _ozu { get; set; }
        /// <summary>
        /// Группа возраста
        /// </summary>
        private int _grvoz { get; set; }
        /// <summary>
        /// Количество подроста
        /// </summary>
        private double _npdr { get; set; }

        public Activitie(ExcelHelper excelHelper, Dictionary<string, int> foundColumnsForSettings, int currentRow)
        {
            this._excelHelper = excelHelper;
            this._foundColumnsForSettings = foundColumnsForSettings;
            this._currentRow = currentRow;
        }

        public void StartSettings()
        {
            var appointActivitieOne = AppointActivities();

            if (!_isAppointActivitieOne)
            {
                return;
            }

            foreach (var column in _foundColumnsForSettings)
            {
                if (column.Key == "PRVB")
                {
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item3.ToString());
                }
                if (column.Key == "MER1")
                {
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item1.ToString());
                }
                if (_isAppointActivitieTwo && column.Key == "MER2")
                {
                    var appointActivitieTwo = AppointActivitieTwo();
                    _excelHelper.Set(column.Value, _currentRow, data: appointActivitieOne.Item2.ToString());
                }
            }
        }
        public void CheckActivities(string column, string cellText)
        {
            if (column == Columns.RequiredColumns[0]) // KATL
            {
                if (int.TryParse(cellText, out int result))
                {
                    this._katl = result;
                }
            }
            if (column == Columns.RequiredColumns[1]) // OZU
            {
                if (int.TryParse(cellText, out int result))
                {
                    this._ozu = result;
                }
            }
            if (column == Columns.RequiredColumns[2]) // GR VOZ
            {
                if (int.TryParse(cellText, out int result))
                {
                    this._grvoz = result;
                }
            }
            if (column == Columns.RequiredColumns[3]) // NPDR
            {
                if (double.TryParse(cellText.Replace(',', '.'), out double result))
                {
                    this._npdr = result;
                }
            }
        }
        private (int activitieOne, int activitieTwo, int prvb) AppointActivities()
        {
            if (this._katl == 80)
            {
                if (this._ozu <= 0)
                {
                    if (this._grvoz == 4 || this._grvoz == 5)
                    {
                        _isAppointActivitieOne = true;

                        int activitieOne = 10;
                        int prvb = 100;
                        int activitieTwo = AppointActivitieTwo();

                        var result = (activitieOne, activitieTwo, prvb);
                        return result;
                    }
                }
            }
            return (0, 0, 0);
        }
        private int AppointActivitieTwo()
        {
            _isAppointActivitieTwo = true;

            if (this._npdr >= 0 && this._npdr < 1)
            {
                return 500;
            }
            else if (this._npdr >= 1 && this._npdr < 1.5)
            {
                return 640;
            }
            else if (this._npdr >= 1.5 && this._npdr < 2.5)
            {
                return 690;
            }
            else if (this._npdr >= 2.5)
            {
                return 660;
            }

            _isAppointActivitieTwo = false;
            return 0;
        }
    }
}

