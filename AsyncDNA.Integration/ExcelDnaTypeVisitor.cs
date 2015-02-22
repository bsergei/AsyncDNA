using System;
using ExcelDna.Integration;

namespace AsyncDNA.Integration
{
    public abstract class ExcelDnaTypeVisitor
    {
        private readonly bool isGenerative_;

        protected ExcelDnaTypeVisitor(bool isGenerative)
        {
            isGenerative_ = isGenerative;
        }

        protected virtual object Visit(object value)
        {
            if (value == null)
                return null;

            var missing = value as ExcelMissing;
            if (missing != null)
                return Visit(missing);

            if (value is ExcelError)
                return Visit((ExcelError)value);

            var reference = value as ExcelReference;
            if (reference != null)
                return Visit(reference);

            var empty = value as ExcelEmpty;
            if (empty != null)
                return Visit(empty);

            var array = value as Array;
            if (array != null)
                return Visit(array);

            return value;
        }

        protected virtual object Visit(ExcelMissing value)
        {
            return value;
        }

        protected virtual object Visit(ExcelError value)
        {
            return value;
        }

        protected virtual object Visit(ExcelEmpty value)
        {
            return value;
        }

        protected virtual object Visit(ExcelReference value)
        {
            object refValue = GetReferenceValue(value);
            return Visit(refValue);
        }

        protected virtual object GetReferenceValue(ExcelReference excelReference)
        {
            return excelReference.GetValue();
        }

        protected virtual object Visit(Array array)
        {
            var rank = array.Rank;

            int[] lengths = new int[rank];
            for (int dimension = 0; dimension < rank; dimension++)
                lengths[dimension] = array.GetLength(dimension);

            Array newArray = isGenerative_
                ? Array.CreateInstance(typeof (object), lengths)
                : null;

            Iterate(0, new int[rank], lengths, array, newArray);

            return newArray ?? array;
        }

        private void Iterate(int dimension, int[] indices, int[] lengths, Array array, Array newArray)
        {
            for (int i = 0; i < lengths[dimension]; i++)
            {
                indices[dimension] = i;

                if (dimension + 1 < lengths.Length)
                {
                    Iterate(dimension + 1, indices, lengths, array, newArray);
                }
                else
                {
                    object value = array.GetValue(indices);
                    var newValue = Visit(value);
                    if (newArray != null)
                        newArray.SetValue(newValue, indices);
                }
            }
        }
    }
}