using System;

namespace AsyncDNA
{
    public class ExceptionHelper
    {
        public static string GetErrorMessage(string functionName, Exception e)
        {
            return String.Format("Function '{0}' was executed with errors. Contact application administrator for assistance.", functionName);
        }
    }
}