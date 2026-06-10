using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;

namespace DocFlow.Core.Helpers
{
    internal static class HtmlHelper
    {
        public static HtmlDocument Load(string html)
        {
            var document = new HtmlDocument();
            document.LoadHtml(html ?? string.Empty);
            return document;
        }

        public static string ExtractText(string html)
        {
            var document = Load(html);
            var nodes = document.DocumentNode.SelectNodes("//h1|//h2|//h3|//p|//li|//td|//th");
            if (nodes == null)
            {
                return HtmlEntity.DeEntitize(document.DocumentNode.InnerText ?? string.Empty).Trim();
            }

            return string.Join("\n\n", nodes.Select(node => HtmlEntity.DeEntitize(node.InnerText).Trim()).Where(text => !string.IsNullOrWhiteSpace(text)));
        }

        public static List<Dictionary<string, string>> ExtractTables(string html)
        {
            var result = new List<Dictionary<string, string>>();
            var document = Load(html);
            var table = document.DocumentNode.SelectSingleNode("//table");
            if (table == null)
            {
                return result;
            }

            var rows = table.SelectNodes(".//tr");
            if (rows == null || rows.Count == 0)
            {
                return result;
            }

            var headerCells = rows[0].SelectNodes("./th|./td");
            var headers = new List<string>();
            if (headerCells != null)
            {
                headers = headerCells
                    .Select(cell => HtmlEntity.DeEntitize(cell.InnerText).Trim())
                    .ToList();
            }

            for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
            {
                var cells = rows[rowIndex].SelectNodes("./td|./th");
                if (cells == null)
                {
                    continue;
                }

                var row = new Dictionary<string, string>();
                for (var column = 0; column < cells.Count; column++)
                {
                    var header = column < headers.Count && !string.IsNullOrWhiteSpace(headers[column]) ? headers[column] : "Column" + (column + 1);
                    row[header] = HtmlEntity.DeEntitize(cells[column].InnerText).Trim();
                }

                result.Add(row);
            }

            return result;
        }
    }
}
