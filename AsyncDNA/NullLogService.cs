using System;
using AsyncDNA.Integration;

namespace AsyncDNA
{
    public class NullLogService : ILogService
    {
        public void WriteError(string logMessage, Exception exception)
        {
        }
    }
}