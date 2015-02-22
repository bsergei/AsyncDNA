using System;
using ExcelDna.Integration;

namespace AsyncDNA
{
    public struct ReferenceFunc : IEquatable<ReferenceFunc>
    {
        public readonly ExcelReference ExcelReference;

        public readonly string FuncName;

        public ReferenceFunc(ExcelReference excelReference, string funcName)
        {
            ExcelReference = excelReference;
            FuncName = funcName;
        }

        public bool Equals(ReferenceFunc other)
        {
            return Equals(ExcelReference, other.ExcelReference) && String.Equals(FuncName, other.FuncName, StringComparison.InvariantCultureIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            return obj is ReferenceFunc && Equals((ReferenceFunc) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((ExcelReference != null ? ExcelReference.GetHashCode() : 0)*397) ^ (FuncName != null ? StringComparer.InvariantCultureIgnoreCase.GetHashCode(FuncName) : 0);
            }
        }
    }
}