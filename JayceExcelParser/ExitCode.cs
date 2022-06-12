namespace JayceExcelParser
{
    internal enum ExitCode : int
    {
        UNKNOWN_ERROR = -1,
        SUCCESS = 0,
        INVALID_EXCEL_DIRECTORY,
        INVALID_OUTPUT_DIRECTORY,
        PROGRAM_EXCEPTION,
    }
}
