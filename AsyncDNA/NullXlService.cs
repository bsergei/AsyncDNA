using AsyncDNA.Integration;
using ExcelDna.Integration;
using Reference = AsyncDNA.Integration.Reference;

namespace AsyncDNA
{
    public class NullXlService : IXlService
    {
        public object Call(Reference reference, CallArguments args)
        {
            return ExcelError.ExcelErrorValue;
        }
    }
}