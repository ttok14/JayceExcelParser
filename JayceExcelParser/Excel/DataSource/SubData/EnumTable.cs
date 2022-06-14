using System;
using System.Collections.Generic;
using System.Text;

namespace JayceExcelParser.Excel.DataSource
{
    class EnumType
    {
        public class Element
        {
            public int value;
            public string name;
            public string comment;
        }

        public string typeName;
        public List<Element> elements = new List<Element>();

        public bool isFlags;
    }
}
