using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    interface IActivitie
    {
        void CheckCells(string column, string cellText);
        void FillCells();
    }
}
