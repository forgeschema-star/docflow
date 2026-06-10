namespace DocFlow.Core.Models
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public string Formula { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public double? FontSize { get; set; }
        public string FontColor { get; set; }
        public string BackgroundColor { get; set; }
        public string BorderColor { get; set; }
        public ExcelHorizontalAlignment Alignment { get; set; }

        public static ExcelCell Text(string value) => new ExcelCell { Value = value };
        public static ExcelCell Header(string value) => new ExcelCell { Value = value, FontBold = true };
        public static ExcelCell Calc(string formula) => new ExcelCell { Formula = formula };
    }

    public enum ExcelHorizontalAlignment { Default, Left, Center, Right }
}
