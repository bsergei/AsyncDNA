using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using AsyncDNA.Integration;
using ExcelDna.Integration;

namespace AsyncDNA
{
    internal class ExcelDnaValuesLogger : ExcelDnaTypeVisitor
    {
        public static string GetString(object excelDanaValues)
        {
            var excelDnaValuesLogger = new ExcelDnaValuesLogger();
            excelDnaValuesLogger.Visit(excelDanaValues);

            return excelDnaValuesLogger.builder_.ToString();
        }

        private readonly StringBuilder builder_;
        private readonly Stack<State> state_;

        private ExcelDnaValuesLogger() : base(false)
        {
            builder_ = new StringBuilder();
            state_ = new Stack<State>();
        }

        private void AppendValue(string s)
        {
            bool needSeparator = false;

            if (state_.Count > 0)
            {
                var state = state_.Peek();
                ArrayState arrayState = state as ArrayState;
                if (arrayState != null)
                {
                    needSeparator = arrayState.ElementsCount > 0;
                    arrayState.ElementsCount++;
                }
            }

            if (needSeparator)
                builder_.Append(", ");

            builder_.Append(s);
        }

        protected override object Visit(ExcelEmpty value)
        {
            AppendValue("#Empty");
            return base.Visit(value);
        }

        protected override object Visit(ExcelError value)
        {
            AppendValue("#Error-" + value);
            return base.Visit(value);
        }

        protected override object Visit(ExcelMissing value)
        {
            AppendValue("#Missing");
            return base.Visit(value);
        }

        protected override object Visit(ExcelReference value)
        {
            AppendValue("#Ref-" + value.GetRange() + "=");
            var refState = new RefState();
            state_.Push(refState);
            
            object visited = base.Visit(value);
            
            state_.Pop();
            return visited;
        }

        protected override object GetReferenceValue(ExcelReference excelReference)
        {
            object refValue;
            try
            {
                refValue = base.GetReferenceValue(excelReference);
            }
            catch (XlCallException e)
            {
                return @"#GetRefValueFailed=" + e.xlReturn;
            }
            catch (Exception e)
            {
                return @"#GetRefValueFailed
~~~~~
" + e + @"
~~~~~";
            }
            return refValue;
        }

        protected override object Visit(Array array)
        {
            AppendValue("".PadRight(array.Rank, '{'));
            state_.Push(new ArrayState());

            var visited = base.Visit(array);

            builder_.Append("".PadRight(array.Rank, '}'));
            state_.Pop();
            return visited;
        }

        protected override object Visit(object value)
        {
            IConvertible cnv = value as IConvertible;
            if (cnv != null)
            {
                bool needQuote = cnv.GetTypeCode() == TypeCode.String;
                string s = cnv.ToString(CultureInfo.InvariantCulture);
                AppendValue(needQuote ? "'" + s + "'" : s);
            }
            else if (value == null)
            {
                AppendValue("#null");
            }

            var visit = base.Visit(value);
            return visit;
        }

        private abstract class State
        {
        }

        private class ArrayState : State
        {
            public int ElementsCount;
        }

        private class RefState : State
        {
        }
    }
}