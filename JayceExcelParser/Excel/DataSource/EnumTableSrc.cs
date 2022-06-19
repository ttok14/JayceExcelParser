using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Common;
using JayceExcelParser.Excel.DataSource;

namespace JayceExcelParser.Excel
{
    class EnumTableSrc
    {
        public Dictionary<string, EnumType> EnumContainer { get; private set; } = new Dictionary<string, EnumType>();

        public bool Add(string identifier, EnumType @enum)
        {
            if (string.IsNullOrEmpty(identifier) || @enum == null || EnumContainer.ContainsKey(identifier))
            {
                // Log.Error
                return false;
            }

            EnumContainer.Add(identifier, @enum);

            return true;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            sb.AppendLine($"===== Enum Info =====");

            foreach (var @enum in EnumContainer)
            {
                sb.AppendLine(" ------------------");
                sb.AppendLine($"Type Name : {@enum.Key}");
                sb.AppendLine($"Use Flags : {@enum.Value.useFlags}");

                sb.AppendLine("-- Element Info Below ");

                foreach (var elem in @enum.Value.elements)
                {
                    sb.AppendLine($"name : {elem.name} , value : {elem.value} , comment : {elem.comment}");
                }
            }

            return sb.ToString();
        }
    }
}
