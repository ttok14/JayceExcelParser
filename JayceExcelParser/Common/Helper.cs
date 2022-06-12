using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using JayceExcelParser.Generate;

namespace JayceExcelParser.Common
{
    static class Helper
    {
        public static ExitCode GenerateResultToExitCode(GenerateResult result)
        {
            if (result == GenerateResult.SUCCESS)
            {
                return ExitCode.SUCCESS;
            }
            else if (result == GenerateResult.FAIL_EXCEL_PATH)
            {
                return ExitCode.INVALID_EXCEL_DIRECTORY;
            }
            else if (result == GenerateResult.FAIL_OUTPUT_FORMAT_NOT_SPECIFIED)
            {
                return ExitCode.INVALID_OUTPUT_DIRECTORY;
            }
            else if (result == GenerateResult.FAIL_EXCEPTION)
            {
                return ExitCode.PROGRAM_EXCEPTION;
            }

            return ExitCode.UNKNOWN_ERROR;
        }

        public static bool IsValidDirectory(string directory)
        {
            if (string.IsNullOrEmpty(directory))
            {
                return false;
            }

            return Directory.Exists(directory);
        }
    }
}
