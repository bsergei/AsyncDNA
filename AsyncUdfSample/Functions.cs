using ExcelDna.Integration;

namespace AsyncUdfSample
{
    public static class Functions
    {
        /// <summary>
        /// Async function sample. Please note it should have IsThreadSafe = false, IsVolatile=false and IsMacroType=true.
        /// </summary>
        [ExcelFunction(Name = AsyncAddRequest.FuncName, Description = "A bit more than your usual adding function.", 
            IsThreadSafe = false, 
            IsVolatile = false, 
            IsMacroType = true)]
        public static object AsyncAdd(object val1, object val2)
        {
            return Addin.Instance.Async.Calc(new AsyncAddRequest(val1, val2));
        }
    }
}
