using System;
using System.Collections.Generic;
using System.Text;

namespace JayceExcelParser.Excel
{
    [Flags]
    internal enum ValueCondition
    {
        None = 0,
        Integer = 0x1,
        ValidIdentifier = 0x1 << 1,
    }

    struct ValueConditionOption
    {
        public ValueCondition condition;

        public ValueConditionOption(ValueCondition condition)
        {
            this.condition = condition;
        }
    }
}
