using System;
using System.IO;
using AsyncDNA;
using ExcelDna.Integration;

namespace AsyncUdfSample
{
    /// <summary>
    /// Sample add-in definition.
    /// </summary>
    public class Addin : IExcelAddIn
    {
        private static Addin instance_;
        private bool initialized_;
        private Async async_;

        public static Addin Instance
        {
            get
            {
                if (instance_ == null
                    || !instance_.initialized_)
                    throw new ApplicationException("Not yet initialized");

                return instance_;
            }
        }

        public Async Async
        {
            get { return async_; }
        }

        public void AutoOpen()
        {
            if (initialized_)
                throw new ApplicationException("Already initialized");

            instance_ = this;
            initialized_ = true;

            async_ = new Async(new XlService(), null);
            InitializeAsyncFuncs();
        }

        private void InitializeAsyncFuncs()
        {
            // All async functions should be registered.
            async_.RegisterAsyncFunc(AsyncAddRequest.FuncName);
        }

        public void AutoClose()
        {
            Dispose();
        }

        public void Dispose()
        {
            async_.Dispose();
            instance_ = null;
        }

        private string GetBinDirectory()
        {
            return Path.GetDirectoryName((string)XlCall.Excel(XlCall.xlGetName));
        }
    }
}