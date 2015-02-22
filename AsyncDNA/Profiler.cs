using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace AsyncDNA
{
    public class Profiler : IDisposable
    {
        private readonly string hint_;
        private readonly bool additive_;
        private readonly Stopwatch sw_ = Stopwatch.StartNew();
        private static readonly Dictionary<string, long> additives_ = new Dictionary<string, long>();

        public Profiler(string hint, bool additive = false)
        {
            hint_ = hint;
            additive_ = additive;
        }

        public void Dispose()
        {
            if (additive_)
            {
                long value;
                additives_.TryGetValue(hint_, out value);
                value += sw_.ElapsedMilliseconds;
                additives_[hint_] = value;

                Debug.WriteLine(hint_ + " " + value);
            }
            else
            {
                Debug.WriteLine(hint_ + " " + sw_.ElapsedMilliseconds);
            }
        }
    }
}