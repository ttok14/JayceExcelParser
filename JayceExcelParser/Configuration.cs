using System;

namespace JayceExcelParser
{
    internal enum Mode
    {
        REAL = 0,
        SIMULATION,
    }

    static class Configuration
    {
        public static class Rules
        {
            /// <summary> Represent the ignorable cell with a prefix </summary>
            public static string IgnoreCasePrefix = "_";

            /// <summary> Represent an enum sheet </summary>
            public static string EnumSheetPrefix = "enum_";

            /// <summary> Represent an schema definition sheet </summary>
            public static string SchemaSheetPrefix = "schema_";
        }

        public static class Simulation
        {
            public static string[] ExcelPaths = new string[]
            {
                @"C:\Users\Jayce\Desktop\Temp\excelTest\tables\Ability_table.xlsx"
            };
        }

        public const Mode CurrentMode = Mode.SIMULATION;

        internal static void Setup(ref string[] args)
        {
            if (CurrentMode == Mode.SIMULATION)
            {
                if (args != null)
                {
                    // args = new string[] { @"-xl:C:\Users\Jayce\Desktop\Temp\excelTest\tables" };
                }
            }
        }
    }
}
