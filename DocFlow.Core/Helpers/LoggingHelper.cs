using DocFlow.Core.Interfaces;
using DocFlow.Core.Models;

namespace DocFlow.Core.Helpers
{
    internal static class LoggingHelper
    {
        public static void LogStart(ILogger logger, DocFlowSettings settings, string operation, string details)
        {
            if (settings != null && !settings.LoggingEnabled)
            {
                return;
            }

            logger.LogInformation("START " + operation + " " + details);
        }

        public static void LogEnd(ILogger logger, DocFlowSettings settings, string operation, string details)
        {
            if (settings != null && !settings.LoggingEnabled)
            {
                return;
            }

            logger.LogInformation("END " + operation + " " + details);
        }
    }
}
