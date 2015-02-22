using System.Threading;
using AsyncDNA.Integration;

namespace AsyncUdfSample
{
    public class AsyncAddRequest : CallArguments, IXlRequest
    {
        public const string FuncName = "AsyncAdd";

        private double P1
        {
            get { return (double) ConvertedArgs[0]; }
        }

        private double P2
        {
            get { return (double) ConvertedArgs[1]; }
        }

        public AsyncAddRequest(object p1, object p2)
            : base(FuncName, p1, p2)
        {
        }

        public object Calc(Reference reference)
        {
            Thread.Sleep(200); // Imitate long work.
            return P1 + P2;
        }
    }
}