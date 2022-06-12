using System;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Common;
using JayceExcelParser.Excel;

namespace JayceExcelParser.Generate
{
    class Generator
    {
        /// <summary>
        /// Generate the output by the specified info
        /// </summary>
        /// <param name="excelDirectory"> An directory where the desired .xlsx files are located (엑셀 파일들 위치한 디렉터리) </param>
        /// <param name="outputDirectory"> An directory in which the output will be saved (결과물들이 저장될 디렉터리) </param>
        /// <returns></returns>
        public GenerateResult Generate(GeneratorDesc desc)
        {
            if (desc == null)
            {
                return GenerateResult.FAIL_EXCEPTION;
            }
            else if (Helper.IsValidDirectory(desc.ExcelDirectory) == false)
            {
                return GenerateResult.FAIL_EXCEL_PATH;
            }
            else if (Helper.IsValidDirectory(desc.OutputDirectory) == false)
            {
                return GenerateResult.FAIL_OUTPUT_PATH;
            }
            else if (desc.IsOutputSpecified == false)
            {
                return GenerateResult.FAIL_OUTPUT_FORMAT_NOT_SPECIFIED;
            }

            var excelReader = new ExcelReader();
            excelReader.Read(desc.ExcelDirectory, out var excelSrc);

            return GenerateResult.SUCCESS;
        }
    }
}
