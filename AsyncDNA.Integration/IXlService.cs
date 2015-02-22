namespace AsyncDNA.Integration
{
    public interface IXlService
    {
        object Call(Reference reference, CallArguments args);
    }
}