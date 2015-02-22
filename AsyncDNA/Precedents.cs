using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelFormulaParser;

namespace AsyncDNA
{
    /// <summary>
    /// Calculates all precedents for specified reference.
    /// </summary>
    public class Precedents
    {
        private readonly ExcelReference src_;

        public Precedents(ExcelReference src)
        {
            src_ = src;
        }

        public IEnumerable<ExcelReference> GetRefs()
        {
            return GetPrecedents(src_).Select(_ => _.Reference).Distinct();
        }

        public IEnumerable<ReferenceFunc> GetFormulas()
        {
            return
                GetPrecedents(src_)
                    .Where(_ => !src_.Equals(_.ParentReference))
                    .Select(_ => new
                    {
                        Token = _.ParentFormula
                            .FirstOrDefault(
                                f =>
                                    f.Type == ExcelFormulaTokenType.Function &&
                                    f.Subtype == ExcelFormulaTokenSubtype.Start),
                        Ref = _.ParentReference
                    })
                    .Where(_ => _.Token != null)
                    .Select(_ => new ReferenceFunc(_.Ref, _.Token.Value))
                    .Distinct();
        }

        private static readonly Regex rangeRegex_ =
            new Regex(
                @"^(?:(?<Sheet>[^!]+)!)?(?:\$?(?<Col1>[a-z]+)\$?(?<Row1>\d+))(?:\:\$?(?<Col2>[a-z]+)\$?(?<Row2>\d+))?$",
                RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly IDictionary<ExcelReference, List<ExcelPrecedent>> cache_ =
            new Dictionary<ExcelReference, List<ExcelPrecedent>>();

        // TODO Handle INDIRECT, etc.
        // TODO Static cache (string-ExcelFormula).
        // TODO Fix for R1C1 mode.
        // TODO Fix external references.
        private IEnumerable<ExcelPrecedent> GetPrecedents(ExcelReference reference)
        {
            List<ExcelPrecedent> items;
            if (cache_.TryGetValue(reference, out items))
            {
                foreach (var item in items)
                {
                    yield return item;
                    foreach (var p in GetPrecedents(item.Reference))
                        yield return p;
                }
            }
            else
            {
                bool isFormula = (bool) XlCall.Excel(XlCall.xlfGetCell, 48, reference);
                if (isFormula)
                {
                    string formula = (string) XlCall.Excel(XlCall.xlfGetCell, 6, reference);
                    ExcelFormula excelFormula = new ExcelFormula(formula);

                    foreach (var token in excelFormula)
                    {
                        if (token.Type == ExcelFormulaTokenType.Operand &&
                            token.Subtype == ExcelFormulaTokenSubtype.Range)
                        {
                            //if (isRCmode)
                            //{
                            //    var regex = new Regex(
                            //        @"^=(?:(?<Sheet>[^!]+)!)?(?:R((?<RAbs>\d+)|(?<RRel>\[-?\d+\]))C((?<CAbs>\d+)|(?<CRel>\[-?\d+\]))){1,2}$",
                            //        RegexOptions.Compiled | RegexOptions.IgnoreCase);
                            //    if (regex.IsMatch(formula))
                            //    {
                            //        throw new NotSupportedException();
                            //    }
                            //}

                            var range = token.Value;
                            Match rangeMatch = rangeRegex_.Match(range);

                            var col1 = ExcelAColumnToInt(rangeMatch.Groups["Col1"].Value) - 1;
                            var row1 = Int32.Parse(rangeMatch.Groups["Row1"].Value) - 1;
                            Group sheetGroup = rangeMatch.Groups["Sheet"];
                            var sheetName = sheetGroup.Success ? sheetGroup.Value : null;

                            int col2 = col1;
                            int row2 = row1;
                            if (rangeMatch.Groups["Col2"].Success)
                            {
                                col2 = ExcelAColumnToInt(rangeMatch.Groups["Col2"].Value) - 1;
                                row2 = Int32.Parse(rangeMatch.Groups["Row2"].Value) - 1;
                            }

                            for (int col = col1; col <= col2; col++)
                            {
                                for (int row = row1; row <= row2; row++)
                                {
                                    ExcelReference precedantRef;

                                    if (sheetName == null)
                                    {
                                        precedantRef = new ExcelReference(row, row, col, col, reference.SheetId);
                                    }
                                    else
                                    {
                                        precedantRef = new ExcelReference(row, row, col, col, sheetName);
                                    }

                                    ExcelPrecedent newPrecedent = new ExcelPrecedent(
                                        precedantRef, excelFormula, formula, reference);

                                    AddToCache(reference, newPrecedent);

                                    yield return newPrecedent;
                                    foreach (var nestedPrecedant in GetPrecedents(precedantRef))
                                        yield return nestedPrecedant;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void AddToCache(ExcelReference reference, ExcelPrecedent newPrecedent)
        {
            List<ExcelPrecedent> cachedItems;
            if (!cache_.TryGetValue(reference, out cachedItems))
            {
                cachedItems = new List<ExcelPrecedent>();
                cache_[reference] = cachedItems;
            }
            cachedItems.Add(newPrecedent);
        }

        private static int ExcelAColumnToInt(string strCol)
        {
            var strColUpperCase = strCol.ToUpper(CultureInfo.InvariantCulture);
            int result = 0;
            int pBase = 1;
            for (int i = strColUpperCase.Length - 1; i >= 0; i--)
            {
                int digit = strColUpperCase[i] - 'A' + 1;
                result += pBase * digit;
                pBase = pBase * 26;
            }
            return result;
        }

        //private static string GetExcelAColumnName(int columnNumber)
        //{
        //    int dividend = columnNumber;
        //    string columnName = String.Empty;
        //    int modulo;

        //    while (dividend > 0)
        //    {
        //        modulo = (dividend - 1) % ('Z' - 'A');
        //        columnName = Convert.ToChar('A' + modulo) + columnName;
        //        dividend = (int)((dividend - modulo) / 26);
        //    }

        //    return columnName;
        //}

        private class ExcelPrecedent
        {
            private readonly ExcelReference reference_;
            private readonly ExcelReference parentReference_;
            private readonly ExcelFormula parentFormula_;
            private readonly string parentFormulaSrc_;

            public ExcelPrecedent(ExcelReference reference, ExcelFormula parentFormula, string parentFormulaSrc, ExcelReference parentReference)
            {
                reference_ = reference;
                parentFormula_ = parentFormula;
                parentFormulaSrc_ = parentFormulaSrc;
                parentReference_ = parentReference;
            }

            public ExcelReference Reference
            {
                get { return reference_; }
            }

            public ExcelFormula ParentFormula
            {
                get { return parentFormula_; }
            }

            public string ParentFormulaSrc
            {
                get { return parentFormulaSrc_; }
            }

            public ExcelReference ParentReference
            {
                get { return parentReference_; }
            }
        }
    }
}