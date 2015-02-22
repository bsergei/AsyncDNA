using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace AsyncDNA.Integration
{
    public class CallArguments
    {
        private readonly string functionName_;
        private readonly object[] rawArgs_;
        private readonly object[] convertedArgs_;

        public CallArguments(string functionName, params object[] rawArgs)
        {
            functionName_ = functionName;
            rawArgs_ = rawArgs;
            convertedArgs_ = new XlTypeConverter().Convert(rawArgs_);
        }

        public string FunctionName
        {
            get { return functionName_; }
        }

        public object[] RawArgs
        {
            get { return rawArgs_; }
        }

        public object[] ConvertedArgs
        {
            get { return convertedArgs_; }
        }
    }
}