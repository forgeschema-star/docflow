namespace DocFlow.Core.Models
{
    public class WordBlock
    {
        public WordBlockType Type { get; set; }
        public string Text { get; set; }
        public int Level { get; set; } = 1;
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public bool Underline { get; set; }
        public double? FontSize { get; set; }
        public string FontColor { get; set; }
        public string BackgroundColor { get; set; }
        public string Alignment { get; set; }

        public static WordBlock Heading(string text, int level = 1) =>
            new WordBlock { Type = WordBlockType.Heading, Text = text, Level = level };

        public static WordBlock Para(string text) =>
            new WordBlock { Type = WordBlockType.Paragraph, Text = text };
    }

    public enum WordBlockType { Paragraph, Heading }
}
