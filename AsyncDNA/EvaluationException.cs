using System;

namespace AsyncDNA
{
    public class EvaluationException : Exception
    {
        public object XllResult { get; private set; }

        public EvaluationException(string message, Exception e, object xllResult)  : base(message, e)
        {
            XllResult = xllResult;
        }
    }
}
