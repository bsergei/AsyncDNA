namespace AsyncDNA.Integration
{
    public class XlTypeConverter : ExcelDnaTypeVisitor
    {
        public XlTypeConverter()
            : base(true)
        {
        }

        public object[] Convert(object[] values)
        {
            return (object[])Convert((object)values);
        }

        public object Convert(object value)
        {
            return Visit(value);
        }
    }
}