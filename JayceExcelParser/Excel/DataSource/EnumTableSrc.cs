using System;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Excel.DataSource;
using Serilog;

namespace JayceExcelParser.Excel
{
    class EnumTableSrc
    {
        public Dictionary<string, EnumType> EnumContainer { get; private set; } = new Dictionary<string, EnumType>();

        public bool Add(string typeName, EnumType @enum)
        {
            if (string.IsNullOrEmpty(typeName) || @enum == null || EnumContainer.ContainsKey(typeName))
            {
                // Log.Error
                return false;
            }

            EnumContainer.Add(typeName, @enum);

            return true;
        }
    }
}
