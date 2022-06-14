using System;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Common;
using JayceExcelParser.Excel;
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

            while ((curRow = FindNextStartRow(sheet, curRow)) != -1)
            {
                Console.WriteLine("======== Found Row : " + curRow);
                Console.WriteLine("TypeName : " + sheet.Cells[curRow, 2].Text);

                Enum Parsing 이어서 진행 ㄱㄱ 
                curRow++;
            }

            return result;
        }

        private int FindNextStartRow(ExcelWorksheet sheet, int curRow)
        {
            return sheet.FindRow(
                1
                , match: (content) =>
                {
                    return content.StartsWith("enum", StringComparison.OrdinalIgnoreCase);
                }, flags: ReadFlags.IgnoreCase
                , beginRow: curRow);
        }
    }
}
