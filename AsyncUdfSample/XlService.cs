using AsyncDNA.Integration;

namespace AsyncUdfSample
{
    public class XlService : IXlService
    {
        public object Call(Reference reference, CallArguments args)
        {
            IXlRequest s = args as IXlRequest;
            if (s != null)
                return s.Calc(reference);

            return null;
        }
    }
}