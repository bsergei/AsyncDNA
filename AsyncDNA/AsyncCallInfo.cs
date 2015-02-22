using System;
using ExcelDna.Integration;

namespace AsyncDNA
{
    // Encapsulates the information that defines and async call or observable hook-up.
    // Checked for equality and stored in a Dictionary, so we have to be careful
    // to define value equality and a consistent HashCode.

    // Used as Keys in a Dictionary - should be immutable. 
    // We allow parameters to be null or primitives or ExcelReference objects, 
    // or a 1D array or 2D array with primitives or arrays.
    internal struct AsyncCallInfo : IEquatable<AsyncCallInfo>
    {
        readonly string _functionName;
        readonly object _parameters;
        readonly int _hashCode;

        public AsyncCallInfo(string functionName, object parameters)
        {
            _functionName = functionName;
            _parameters = parameters;
            _hashCode = 0; // Need to set to some value before we call a method.
            _hashCode = ComputeHashCode();
        }

        // Jon Skeet: http://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode
        int ComputeHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + (_functionName == null ? 0 : _functionName.GetHashCode());
                hash = hash * 23 + ComputeHashCode(_parameters);
                return hash;
            }
        }

        // Computes a hash code for the parameters, consistent with the value equality that we define.
        // Also checks that the data types passed for parameters are among those we handle properly for value equality.
        // For now no string[]. But invalid types passed in will causes exception immediately.
        static int ComputeHashCode(object obj)
        {
            if (obj == null) return 0;

            // CONSIDER: All of this could be replaced by a check for (obj is ValueType || obj is ExcelReference)
            //           which would allow a more flexible set of parameters, at the risk of comparisons going wrong.
            //           We can reconsider if this arises, or when we implement async automatically or custom marshaling 
            //           to other data types. For now this allow everything that can be passed as parameters from Excel-DNA.
            if (obj is double ||
                obj is float ||
                obj is string ||
                obj is bool ||
                obj is DateTime ||
                obj is ExcelReference ||
                obj is ExcelError ||
                obj is ExcelEmpty ||
                obj is ExcelMissing ||
                obj is int ||
                obj is uint ||
                obj is long ||
                obj is ulong ||
                obj is short ||
                obj is ushort ||
                obj is byte ||
                obj is sbyte ||
                obj is decimal ||
                obj.GetType().IsEnum)
            {
                return obj.GetHashCode();
            }

            unchecked
            {
                int hash = 17;

                double[] doubles = obj as double[];
                if (doubles != null)
                {
                    foreach (double item in doubles)
                    {
                        hash = hash * 23 + item.GetHashCode();
                    }
                    return hash;
                }

                double[,] doubles2 = obj as double[,];
                if (doubles2 != null)
                {
                    foreach (double item in doubles2)
                    {
                        hash = hash * 23 + item.GetHashCode();
                    }
                    return hash;
                }

                object[] objects = obj as object[];
                if (objects != null)
                {
                    foreach (object item in objects)
                    {
                        hash = hash * 23 + ((item == null) ? 0 : ComputeHashCode(item));
                    }
                    return hash;
                }

                object[,] objects2 = obj as object[,];
                if (objects2 != null)
                {
                    foreach (object item in objects2)
                    {
                        hash = hash * 23 + ((item == null) ? 0 : ComputeHashCode(item));
                    }
                    return hash;
                }
            }
            throw new ArgumentException("Invalid type used for async parameter(s)", "parameters");
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (obj.GetType() != typeof(AsyncCallInfo)) return false;
            return Equals((AsyncCallInfo)obj);
        }

        public bool Equals(AsyncCallInfo other)
        {
            if (_hashCode != other._hashCode) return false;
            return Equals(other._functionName, _functionName)
                   && ValueEquals(_parameters, other._parameters);
        }

        #region Helpers to implement value equality
        // The value equality we check here is for the types we allow in CheckParameterTypes above.
        static bool ValueEquals(object a, object b)
        {
            if (Equals(a, b)) return true; // Includes check for both null
            if (a is double[] && b is double[]) return ArrayEquals((double[])a, (double[])b);
            if (a is double[,] && b is double[,]) return ArrayEquals((double[,])a, (double[,])b);
            if (a is object[] && b is object[]) return ArrayEquals((object[])a, (object[])b);
            if (a is object[,] && b is object[,]) return ArrayEquals((object[,])a, (object[,])b);
            return false;
        }

        static bool ArrayEquals(double[] a, double[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }

        static bool ArrayEquals(double[,] a, double[,] b)
        {
            int rows = a.GetLength(0);
            int cols = a.GetLength(1);
            if (rows != b.GetLength(0) ||
                cols != b.GetLength(1))
            {
                return false;
            }
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (a[i, j] != b[i, j]) return false;
                }
            }
            return true;
        }

        static bool ArrayEquals(object[] a, object[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (!ValueEquals(a[i], b[i]))
                    return false;
            }
            return true;
        }

        static bool ArrayEquals(object[,] a, object[,] b)
        {
            int rows = a.GetLength(0);
            int cols = a.GetLength(1);
            if (rows != b.GetLength(0) ||
                cols != b.GetLength(1))
            {
                return false;
            }
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (!ValueEquals(a[i, j], b[i, j]))
                        return false;
                }
            }
            return true;
        }
        #endregion

        public override int GetHashCode()
        {
            return _hashCode;
        }
    }
}