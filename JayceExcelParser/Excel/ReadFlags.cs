using System;
using System.Collections.Generic;
using System.Text;

namespace JayceExcelParser.Excel
{
    [System.Flags]
    internal enum ReadFlags
    {
        None = 0x00,
        IgnoreCase = 0x01,
    }
}
