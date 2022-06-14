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

        public static bool ForeachRow(this ExcelWorksheet sheet, int column, ReadFlags flags, Action<ExcelAddress, string> onRead)
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

            var startRow = dms.Start.Row;
            var endRow = dms.End.Row;

            for (int row = startRow; row <= endRow; row++)
            {
                var cell = sheet.Cells[row, column];

                if (EvaluateFlag(flags, cell.Text))
                {
                    onRead.Invoke(cell, cell.Text);
                }
            }

            return true;
        }

        public static int FindColumn(this ExcelWorksheet sheet, int row, Predicate<string> match, ReadFlags flags = ReadFlags.None, int beginColumn = 1, int endColumn = -1)
        {
            if (match == null)
            {
                return -1;
            }

            var dms = sheet.Dimension;
            if (dms == null)
            {
                return -1;
            }

            beginColumn = sheet.GetClampedColumn(beginColumn);

            if (endColumn < 1)
            {
                endColumn = dms.End.Column;
            }
            else
            {
                endColumn = sheet.GetClampedColumn(endColumn);
            }

            int result = -1;

            for (int column = beginColumn; column <= endColumn; column++)
            {
                var cell = sheet.Cells[row, column];

                if (EvaluateFlag(flags, cell.Text) && match.Invoke(cell.Text))
                {
                    result = column;
                    break;
                }
            }

            return result;
        }

        public static int FindRow(this ExcelWorksheet sheet, int column, Predicate<string> match, ReadFlags flags = ReadFlags.None, int beginRow = 1, int endRow = -1)
        {
            if (match == null)
            {
                return -1;
            }

            var dms = sheet.Dimension;
            if (dms == null)
            {
                return -1;
            }

            beginRow = sheet.GetClampedRow(beginRow);

            if (endRow < 1)
            {
                endRow = dms.End.Row;
            }
            else
            {
                endRow = sheet.GetClampedRow(endRow);
            }

            int result = -1;

            for (int row = beginRow; row <= endRow; row++)
            {
                var cell = sheet.Cells[row, column];

                if (EvaluateFlag(flags, cell.Text) && match.Invoke(cell.Text))
                {
                    result = row;
                    break;
                }
            }

            return result;
        }

        public static int GetClampedColumn(this ExcelWorksheet sheet, int desiredColumn)
        {
            var dms = sheet.Dimension;
            if (dms == null)
            {
                return 1;
            }

            if (desiredColumn < 1)
            {
                return 1;
            }
            else if (desiredColumn > dms.End.Column)
            {
                return dms.End.Column;
            }

            return desiredColumn;
        }

        public static int GetClampedRow(this ExcelWorksheet sheet, int desiredRow)
        {
            var dms = sheet.Dimension;
            if (dms == null)
            {
                return 1;
            }

            if (desiredRow < 1)
            {
                return 1;
            }
            else if (desiredRow > dms.End.Row)
            {
                return dms.End.Row;
            }

            return desiredRow;
        }

        public static bool ToSheetType(string sheetName, out SheetType resultSheetType)
        {
            resultSheetType = default(SheetType);

            if (string.IsNullOrEmpty(sheetName))
            {
                return false;
            }

            if (sheetName.Equals(Configuration.Rules.EnumSheetPredefinedName, StringComparison.OrdinalIgnoreCase))
            {
                resultSheetType = SheetType.Enum;
            }
            else if (sheetName.StartsWith(Configuration.Rules.SchemaSheetPrefix, StringComparison.OrdinalIgnoreCase))
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
            //if (string.IsNullOrEmpty(value))
            //{
            //    return false;
            //}

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
