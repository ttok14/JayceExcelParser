using System;
using System.Collections.Generic;
using System.Text;

namespace JayceExcelParser.Excel
{
    class ExcelSrc
    {
        public EnumTableSrc Enum { get; private set; }
        public ExcelContentSrc Content { get; private set; }
        public SchemaSrc SchemaInfo { get; private set; }

        public void SetEnum(EnumTableSrc @enum)
        {
            this.Enum = @enum;
        }
    }
}
