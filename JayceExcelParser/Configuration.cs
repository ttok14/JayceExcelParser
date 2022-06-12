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
