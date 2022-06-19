using Serilog;

namespace JayceExcelParser.Common
{
    static class JLog
    {
        public static void Information(string msg)
        {
            if (IsEnabled() == false)
            {
                return;
            }

            Log.Information(msg);
        }

        public static void Warning(string msg)
        {
            if (IsEnabled() == false)
            {
                return;
            }

            Log.Warning(msg);
        }

        public static void Error(string msg)
        {
            if (IsEnabled() == false)
            {
                return;
            }

            Log.Error(msg);
        }

        public static bool IsEnabled()
        {
            return Configuration.ENABLE_LOG;
        }
    }
}
