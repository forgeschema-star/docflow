using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocFlow.Core.Helpers
{
    internal static class TextStructureHelper
    {
        private static readonly Regex TableSeparatorRegex = new Regex(@"\s{3,}|\t+|\|", RegexOptions.Compiled);

        public static IList<TextBlock> ParseBlocks(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return new List<TextBlock>();
            }

            var normalized = content.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalized.Split('\n');
            var blocks = new List<TextBlock>();
            var paragraphLines = new List<string>();

            foreach (var rawLine in lines)
            {
                var line = rawLine == null ? string.Empty : rawLine.TrimEnd();
                if (string.IsNullOrWhiteSpace(line))
                {
                    FlushParagraph(paragraphLines, blocks);
                    continue;
                }

                var headingLevel = InferHeadingLevel(line);
                if (headingLevel > 0)
                {
                    FlushParagraph(paragraphLines, blocks);
                    blocks.Add(new TextBlock(TextBlockType.Heading, NormalizeHeadingText(line), headingLevel));
                    continue;
                }

                paragraphLines.Add(line.Trim());
            }

            FlushParagraph(paragraphLines, blocks);
            return blocks;
        }

        public static bool TrySplitTableRow(string line, out List<string> cells)
        {
            cells = new List<string>();
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            var parts = TableSeparatorRegex
                .Split(line.Trim())
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .Select(part => part.Trim())
                .ToList();

            if (parts.Count < 2)
            {
                return false;
            }

            cells = parts;
            return true;
        }

        private static void FlushParagraph(List<string> paragraphLines, ICollection<TextBlock> blocks)
        {
            if (paragraphLines.Count == 0)
            {
                return;
            }

            blocks.Add(new TextBlock(
                TextBlockType.Paragraph,
                string.Join(" ", paragraphLines.Where(line => !string.IsNullOrWhiteSpace(line))),
                0));

            paragraphLines.Clear();
        }

        private static int InferHeadingLevel(string line)
        {
            if (line.StartsWith("### ", StringComparison.Ordinal))
            {
                return 3;
            }

            if (line.StartsWith("## ", StringComparison.Ordinal))
            {
                return 2;
            }

            if (line.StartsWith("# ", StringComparison.Ordinal))
            {
                return 1;
            }

            var trimmed = line.Trim();
            if (trimmed.Length <= 80 &&
                !trimmed.EndsWith(".", StringComparison.Ordinal) &&
                !trimmed.EndsWith(";", StringComparison.Ordinal) &&
                (trimmed.Equals(trimmed.ToUpperInvariant(), StringComparison.Ordinal) || trimmed.EndsWith(":", StringComparison.Ordinal)))
            {
                return 1;
            }

            return 0;
        }

        private static string NormalizeHeadingText(string line)
        {
            return line.TrimStart('#', ' ').Trim();
        }
    }

    internal enum TextBlockType
    {
        Paragraph = 0,
        Heading = 1
    }

    internal sealed class TextBlock
    {
        public TextBlock(TextBlockType type, string text, int level)
        {
            Type = type;
            Text = text ?? string.Empty;
            Level = level;
        }

        public TextBlockType Type { get; private set; }

        public string Text { get; private set; }

        public int Level { get; private set; }
    }
}
