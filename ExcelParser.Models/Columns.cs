using System.Collections.Generic;

namespace ExcelParser.Models
{
    public static class Columns
    {
        public static List<string> RequiredColumns = new List<string>
        {   "KATL",
            "OZU",
            "Gr_voz",
            "NPDR"
        };
        public static List<string> ColumnsForSettings = new List<string>
        {
            "PRVB",
            "MER1",
            "MER2"
        };
    }
}
