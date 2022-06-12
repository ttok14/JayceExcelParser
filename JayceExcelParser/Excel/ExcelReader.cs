using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Generate;
using OfficeOpenXml;
using Serilog;

namespace JayceExcelParser.Excel
{
    internal class ExcelReader
    {
        public ExcelReadResult Read(string excelRootDirectory, out ExcelSrc resultExcelSrc)
        {
            resultExcelSrc = new ExcelSrc();

            if (Configuration.CurrentMode == Mode.REAL)
            {
                foreach (string item in Directory.EnumerateFiles(excelRootDirectory, "*.xlsx"))
                {
                    Console.WriteLine($"File : {item}");
                }
            }
            else if (Configuration.CurrentMode == Mode.SIMULATION)
            {
                var excel = new ExcelPackage(new FileInfo(Configuration.Simulation.ExcelPaths[0]));

                foreach (var sheet in excel.Workbook.Worksheets)
                {
                    Console.WriteLine($"Sheet Name : {sheet.Name}");
                }
            }

            return ExcelReadResult.SUCCESS;
        }
    }
}
