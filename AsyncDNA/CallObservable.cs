using System;
using System.Threading;
using AsyncDNA.Integration;

namespace AsyncDNA
{
    public class CallObservable : IObservable<object>, IDisposable
    {
        private readonly Reference reference_;
        private readonly CallArguments args_;
        private volatile bool finished_;

        private readonly IXlService xlService_;
        private readonly ILogService logService_;

        public CallObservable(Reference reference, CallArguments args, IXlService xlService, ILogService logService)
        {
            reference_ = reference;
            args_ = args;
            xlService_ = xlService;
            logService_ = logService;
        }

        public bool Finished
        {
            get { return finished_; }
        }

        public IDisposable Subscribe(IObserver<object> observer)
        {
            ThreadPool.QueueUserWorkItem(s =>
            {
                //Thread.Sleep(2500);
                //observer.OnNext("FF");
                //finished_ = true;
                //observer.OnCompleted();
                try
                {
                    var result = DoCall();
                    observer.OnNext(result);
                    observer.OnCompleted();
                }
                catch (Exception e)
                {
                    observer.OnError(e);
                }
            });
            return this;
        }

        public void Dispose()
        {
        }

        private object DoCall()
        {
            try
            {
                object callResult = xlService_.Call(reference_, args_);
                return callResult;
            }
            catch (EvaluationException e)
            {
                var logMessage = "Error at: " + FunctionLoggerHelper.GetReferenceString(reference_) + @"
converted params: " + FunctionLoggerHelper.LogParameters(args_.FunctionName, args_);

                logService_.WriteError(logMessage, e);

                return e.XllResult;
            }
            catch (Exception e)
            {
                var logMessage = "Error at: " + FunctionLoggerHelper.GetReferenceString(reference_) + @"
converted params: " + FunctionLoggerHelper.LogParameters(args_.FunctionName, args_);

                logService_.WriteError(logMessage, e);
                var message = ExceptionHelper.GetErrorMessage(args_.FunctionName, e);
                return message;
            }
        }
    }
}