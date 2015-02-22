using System;
using AsyncDNA.Integration;

namespace AsyncDNA
{
    public class FunctionLoggerHelper
    {
        public static string LogParameters(string functionName, CallArguments args)
        {
            try
            {
                return String.Format("[functionName='{0}', args='{1}']",
                    functionName ?? "null",
                    args == null ? "null" : args.ToString());
            }
            catch (Exception e)
            {
                return @"[Error logging parameters: 
~~~~~~~~
" + e + @"
~~~~~~~~
]";
            }
        }

        public static string GetReferenceString(Reference reference)
        {
            try
            {
                if (reference == null)
                    return "[Reference is not available]";

                if (reference == Reference.Empty)
                    return "[Reference is empty]";

                return
                    String.Format(@"[Workbook='{0}', sheet='{1}', range='{2}']",
                        reference.Workbook ?? "", reference.Worksheet ?? "", reference.Range);
            }
            catch (Exception e)
            {
                return @"[Failed to get reference
~~~~~~~~
" + e + @"
~~~~~~~~
]";
            }
        }
    }
}