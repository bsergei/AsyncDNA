using System;
using Excel;
using ExcelDna.Integration;
using Reference = AsyncDNA.Integration.Reference;

namespace AsyncDNA
{
    public static class ExcelReferenceExtensions
    {
        public static Range ToComRange(this ExcelReference reference)
        {
            var isRCmode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
            var refText = (string)XlCall.Excel(XlCall.xlfReftext, reference, !isRCmode);
            Range range = ((Excel.Application)ExcelDnaUtil.Application).Range[refText];
            return range;
        }

        public static string GetRange(this ExcelReference reference)
        {
            return String.Format("R{0}C{1}:R{2}C{3}",
                reference.RowFirst + 1, reference.ColumnFirst + 1,
                reference.RowLast + 1, reference.ColumnLast + 1);
        }

        public static Reference XlfToReference(this ExcelReference excelReference)
        {
            string workbookName = (string)XlCall.Excel(XlCall.xlfGetDocument, 68);
            string sheetName = XlCall.Excel(XlCall.xlSheetNm, excelReference) as String;

            return new Reference(workbookName, sheetName, excelReference.SheetId, excelReference.RowFirst, excelReference.RowLast,
                excelReference.ColumnFirst, excelReference.ColumnLast);
        }

        public static Precedents GetPrecedents(this ExcelReference reference)
        {
            return new Precedents(reference);
        }
    }
}