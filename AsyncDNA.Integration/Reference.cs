using System;

namespace AsyncDNA.Integration
{
    public class Reference
    {
        public readonly static Reference Empty = new Reference(null, null, IntPtr.Zero, -1, -1, -1, -1);

        public Reference(string workbook, string worksheet, IntPtr worksheetId, int rowFirst, int rowLast, int columnFirst, int columnLast)
        {
            WorksheetId = worksheetId;
            Workbook = workbook;
            Worksheet = worksheet;
            RowFirst = rowFirst;
            RowLast = rowLast;
            ColumnFirst = columnFirst;
            ColumnLast = columnLast;
        }


        public int RowFirst { get; private set; }

        public int RowLast { get; private set; }

        public int ColumnFirst { get; private set; }

        public int ColumnLast { get; private set; }

        public string Workbook { get; private set; }
        
        public string Worksheet { get; private set; }

        public IntPtr WorksheetId { get; private set; }

        public string Range
        {
            get
            {
                return string.Format("R{0}C{1}:R{2}C{3}",
                    RowFirst + 1, ColumnFirst + 1,
                    RowLast + 1, ColumnLast + 1);
            }
        }
    }
}