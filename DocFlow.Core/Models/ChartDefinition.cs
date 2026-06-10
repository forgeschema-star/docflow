namespace DocFlow.Core.Models
{
    public class ChartDefinition
    {
        public ChartType ChartType { get; set; } = ChartType.Bar;
        public string Title { get; set; } = string.Empty;
        public int HeaderRow { get; set; } = 1;
        public int DataStartRow { get; set; } = 2;
        public int LabelColumn { get; set; } = 1;
        public int FirstDataColumn { get; set; } = 2;
        public int LastDataColumn { get; set; } = -1;
        public int ChartStartColumn { get; set; } = 5;
        public int ChartStartRow { get; set; } = 2;
        public int ChartWidthCells { get; set; } = 8;
        public int ChartHeightCells { get; set; } = 15;
    }

    public enum ChartType { Bar, Line, Pie }
}
