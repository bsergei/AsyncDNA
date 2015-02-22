using System;

namespace AsyncDNA.Integration
{
    public interface ILogService
    {
        void WriteError(string logMessage, Exception exception);
    }
}