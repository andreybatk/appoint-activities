using System.Collections.Generic;

namespace ExcelParser.Models
{
    public static class Columns
    {
        public static List<string> RequiredColumns = new List<string>
        {   "KATL",
            "OZU",
            "Gr_voz",
            "NPDR",
            "POR1",
            "POR2",
            "KS1",
            "KS2",
            "TLU",
            "D1"
        };
        public static List<string> RequiredPolColumns = new List<string>
        {   "POL1",
            "POL2",
            "POL3",
            "POL4",
            "POL5",
            "POL6",
            "POL7",
            "POL8",
            "POL9",
            "POL10"
        };
        public static List<string> RequiredJrColumns = new List<string>
        {   "JR1",
            "JR2",
            "JR3",
            "JR4",
            "JR5",
            "JR6",
            "JR7",
            "JR8",
            "JR9",
            "JR10"
        };
        public static List<string> ColumnsForFilling = new List<string>
        {
            "PRVB",
            "MER1",
            "MER2"
        };
    }
}
