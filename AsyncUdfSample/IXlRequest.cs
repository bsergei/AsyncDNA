using AsyncDNA.Integration;

namespace AsyncUdfSample
{
    public interface IXlRequest
    {
        object Calc(Reference reference);
    }
}