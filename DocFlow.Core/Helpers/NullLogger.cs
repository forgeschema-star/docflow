using System;
using DocFlow.Core.Interfaces;

namespace DocFlow.Core.Helpers
{
    public sealed class NullLogger : ILogger
    {
        public void LogDebug(string message)
        {
        }

        public void LogInformation(string message)
        {
        }

        public void LogWarning(string message)
        {
        }

        public void LogError(string message, Exception exception = null)
        {
        }
    }
}
