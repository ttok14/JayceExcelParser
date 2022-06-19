using System;
using System.Collections.Generic;
using System.Text;

namespace JayceExcelParser.Excel
{
    class EnumElemenReadContext
    {
        public string TypeName;

        public HashSet<string> AddedElementNames = new HashSet<string>();
        public HashSet<int> AddedValues = new HashSet<int>();

        public void Setup(string typeName)
        {
            this.TypeName = typeName;
        }

        public void Add(string elementName, int value)
        {
            AddedElementNames.Add(elementName);
            AddedValues.Add(value);
        }

        public void Clear()
        {
            TypeName = string.Empty;
            AddedElementNames.Clear();
            AddedValues.Clear();
        }
    }
}
