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

            public Element(int value, string name, string comment)
            {
                this.value = value;
                this.name = name;
                this.comment = comment;
            }
        }

        public string identifier;
        public List<Element> elements = new List<Element>();

        public bool useFlags;
    }
}
