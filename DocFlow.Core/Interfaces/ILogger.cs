using System;

namespace DocFlow.Core.Interfaces
{
    public interface ILogger
    {
        void LogDebug(string message);

        void LogInformation(string message);

        void LogWarning(string message);

        void LogError(string message, Exception exception = null);
    }
}
