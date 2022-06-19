using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Generate;
using OfficeOpenXml;
using Serilog;
using JayceExcelParser.Common;

namespace JayceExcelParser.Excel
{
    internal class ExcelReader
    {
        public ExcelReadResult Read(string excelRootDirectory, out ExcelSrc resultExcelSrc)
        {
            if (Configuration.CurrentMode == Mode.SIMULATION)
            {
                resultExcelSrc = new ExcelSrc();
                var excel = new ExcelPackage(new FileInfo(Configuration.Simulation.ExcelPaths[0]));

                foreach (var sheet in excel.Workbook.Worksheets)
                {
                    var dms = sheet.Dimension;
                    ExcelHelper.ToSheetType(sheet.Name, out var sheetType);

                    if (sheetType == SheetType.Enum)
                    {
                        var reader = new EnumReader();
                        var enumOutput = reader.Read(sheet);
                        resultExcelSrc.SetEnum(enumOutput);

                        Console.WriteLine(enumOutput.ToString());
                    }

                    if (dms != null)
                    {
                        //Console.WriteLine($"columns : {dms.Columns} , rows;  {dms.Rows}");
                        //var start = dms.Start;
                        //var end = dms.End;

                        //Console.WriteLine($"Start : {start.Address} , {start.ToString()} , {end.ToString()}");
                        //Console.WriteLine($"End Col, Row : {end.Column} , {end.Row}");
                    }

                    // sheet.ForeachColumn(1, ReadFlags.IgnoreCase, (addr, content) => { Console.WriteLine(content); });

                    Console.WriteLine();
                }

                return ExcelReadResult.SUCCESS;
            }

            resultExcelSrc = new ExcelSrc();

            foreach (string item in Directory.EnumerateFiles(excelRootDirectory, "*.xlsx"))
            {
                Console.WriteLine($"File : {item}");
            }

            return ExcelReadResult.SUCCESS;
        }
    }
}
