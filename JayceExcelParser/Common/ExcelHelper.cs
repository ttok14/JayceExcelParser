using System;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Excel;
using OfficeOpenXml;

namespace JayceExcelParser.Common
{
    static class ExcelHelper
    {
        #region ======:: ExcelWorkSheet Extension ::=======

        public static string GetValueEx(this ExcelWorksheet sheet, int row, int col)
        {
            var dms = sheet.Dimension;
            if (dms == null)
            {
                return string.Empty;
            }

            if (dms.IsRowInRange(row) == false || dms.IsColumnInRange(col) == false)
            {
                return string.Empty;
            }

            return sheet.Cells[row, col]?.Text;
        }

        public static bool IsRowInRange(this ExcelAddressBase dimension, int row)
        {
            if (dimension == null)
            {
                return false;
            }

            return row >= dimension.Start.Row && row <= dimension.End.Row;
        }

        public static bool IsColumnInRange(this ExcelAddressBase dimension, int column)
        {
            if (dimension == null)
            {
                return false;
            }

            return column >= dimension.Start.Column && column <= dimension.End.Column;
        }

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

        public static int FindRow(this ExcelWorksheet sheet, int column, Func<int, string, bool> match, ReadFlags flags = ReadFlags.None, int beginRow = 1, int endRow = -1)
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

                if (EvaluateFlag(flags, cell.Text) && match.Invoke(row, cell.Text))
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

        public static bool InspectValue(this ExcelWorksheet sheet, int col, int row, ValueConditionOption condition)
        {
            return TakeValue(sheet, col, row, condition, out var dummy);
        }

        public static bool TakeValue(this ExcelWorksheet sheet, int col, int row, ValueConditionOption condition, out string cellValue)
        {
            cellValue = GetValueEx(sheet, row, col);

            // Integer 정수 여부
            if (HasFlag((int)condition.condition, (int)ValueCondition.Integer))
            {
                if (int.TryParse(cellValue, out var dummy) == false)
                {
                    return false;
                }
            }

            // 유효한 Identifier 이름 여부
            if (HasFlag((int)condition.condition, (int)ValueCondition.ValidIdentifier))
            {
                if (LexicalHelper.IsValidIdentifierName(cellValue) == false)
                {
                    return false;
                }
            }

            return true;
        }

        #endregion

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

        public static bool EnumHasFlags(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return false;
            }

            return content.Contains("[Flags]", StringComparison.OrdinalIgnoreCase);
        }

        public static bool EvaluateFlag(ReadFlags flag, string value)
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

        public static bool HasFlag(int combined, int targetFlag)
        {
            return (combined & targetFlag) != 0;
        }
    }
}
