using System;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Common;
using JayceExcelParser.Excel;
using JayceExcelParser.Excel.DataSource;
using OfficeOpenXml;

namespace JayceExcelParser.Excel
{
    class EnumReader
    {
        public EnumTableSrc Read(ExcelWorksheet sheet)
        {
            if (ExcelHelper.ToSheetType(sheet.Name, out var sheetType) == false || sheetType != SheetType.Enum)
            {
                return null;
            }

            var dms = sheet.Dimension;
            if (dms == null)
            {
                return null;
            }

            var result = new EnumTableSrc();
            int curRow = 1;

            // enum 시작 Row 를 찾음
            while ((curRow = FindNextStartRow(sheet, curRow)) != -1)
            {
                var enumInstance = new EnumType();

                ReadDefinition(enumInstance, sheet, curRow);
                curRow = ReadElements(enumInstance, enumInstance.identifier, sheet, curRow + 1);

                result.EnumContainer.Add(enumInstance.identifier, enumInstance);
            }

            return result;
        }

        void ReadDefinition(EnumType result, ExcelWorksheet sheet, int row)
        {
            var header = sheet.GetValueEx(row, 1);
            var typeName = sheet.GetValueEx(row, 2);

            result.useFlags = ExcelHelper.EnumHasFlags(header);

            if (Helper.HasWhiteSpace(typeName))
            {
                JLog.Warning($"Type [{typeName}] has one or more whitespaces");
                result.identifier = typeName.Trim();
            }
            else
            {
                result.identifier = typeName;
            }
        }

        EnumElemenReadContext elementReadContext = new EnumElemenReadContext();
        /// <returns> Next Row </returns>
        int ReadElements(EnumType result, string enumIdentifier, ExcelWorksheet sheet, int startRow)
        {
            elementReadContext.Clear();
            elementReadContext.Setup(enumIdentifier);

            // element read 시작 
            bool keepParsing = true;
            int curRow = startRow;
            
            while (keepParsing)
            {
                var element = ParseElement(sheet, curRow, elementReadContext);

                if (element != null)
                {
                    result.elements.Add(element);
                    curRow++;
                }
                else
                {
                    keepParsing = false;
                }
            }

            return curRow + 1;
        }

        EnumType.Element ParseElement(ExcelWorksheet sheet, int row, EnumElemenReadContext context)
        {
            bool isValidInteger = sheet.TakeValue(1, row, new ValueConditionOption(ValueCondition.Integer), out var strElemNum);
            if (isValidInteger == false)
            {
                return null;
            }

            int elemNum = int.Parse(strElemNum);

            // 중복 Add Error
            if (context.AddedValues.Contains(elemNum))
            {
                JLog.Error($"Enum Type [{context.TypeName}] already has an element that has the following value : {elemNum}");
                return null;
            }

            bool isValidName = sheet.TakeValue(2, row, new ValueConditionOption(ValueCondition.ValidIdentifier), out var elemName);
            if (isValidName == false)
            {
                return null;
            }

            // 중복 Element Name Add Error
            if (context.AddedElementNames.Contains(elemName))
            {
                JLog.Error($"Enum Type [{context.TypeName}] already has an element that has the following name : {elemName}");
                return null;
            }

            string comment = sheet.GetValueEx(row, 3);

            return new EnumType.Element(elemNum, elemName, comment);
        }

        private int FindNextStartRow(ExcelWorksheet sheet, int curRow)
        {
            return sheet.FindRow(
                1
                , match: (row, content) =>
                {
                    // Valid 한 definition 인지 체크
                    return content.StartsWith("enum", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(sheet.GetValueEx(row, 2)) == false;
                }, flags: ReadFlags.IgnoreCase
                , beginRow: curRow);
        }
    }
}
