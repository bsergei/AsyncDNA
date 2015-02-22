using System;
using ExcelDna.Integration;

namespace AsyncDNA
{
    public class ExcelObservable<T> : IExcelObservable
    {
        readonly IObservable<T> observable_;

        public ExcelObservable(IObservable<T> observable)
        {
            observable_ = observable;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            return observable_.Subscribe(new ExcelObserverImpl(observer));
        }

        private class ExcelObserverImpl : IObserver<T>
        {
            private readonly IExcelObserver excelObserver_;

            public ExcelObserverImpl(IExcelObserver excelObserver)
            {
                excelObserver_ = excelObserver;
            }

            public void OnNext(T value)
            {
                excelObserver_.OnNext(value);
            }

            public void OnError(Exception error)
            {
                excelObserver_.OnError(error);
            }

            public void OnCompleted()
            {
                excelObserver_.OnCompleted();
            }
        }
    }
}