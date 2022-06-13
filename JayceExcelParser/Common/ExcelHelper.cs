using System;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Excel;
using OfficeOpenXml;

namespace JayceExcelParser.Common
{
    static class ExcelHelper
    {
        public static bool ForeachColumn(this ExcelWorksheet sheet, int row, ReadFlags flags, Action<ExcelAddress, string> onRead)
        {
            if (onRead == null)
            {
                return false;
            }

            var dms = sheet.Dimension;
            if (dms == null)
            {
                return false;
            }

            var startCol = dms.Start.Column;
            var endCol = dms.End.Column;

            for (int col = startCol; col <= endCol; col++)
            {
                var cell = sheet.Cells[row, col];

                if (EvaluateFlag(flags, cell.Text))
                {
                    onRead.Invoke(cell, cell.Text);
                }
            }

            return true;
        }

        public static bool ToSheetType(string sheetName, out SheetType resultSheetType)
        {
            resultSheetType = default(SheetType);

            if (string.IsNullOrEmpty(sheetName))
            {
                return false;
            }

            if (sheetName.StartsWith(Configuration.Rules.EnumSheetPrefix))
            {
                resultSheetType = SheetType.Enum;
            }
            else if (sheetName.StartsWith(Configuration.Rules.SchemaSheetPrefix))
            {
                resultSheetType = SheetType.Schema;
            }
            else
            {
                resultSheetType = SheetType.Content;
            }

            return true;
        }

        static bool EvaluateFlag(ReadFlags flag, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            if (HasFlag((int)flag, (int)ReadFlags.IgnoreCase))
            {
                if (value.StartsWith(Configuration.Rules.IgnoreCasePrefix))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool HasFlag(int combined, int targetFlag)
        {
            return (combined & targetFlag) != 0;
        }
    }
}
