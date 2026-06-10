using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocFlow.Core.Helpers
{
    internal static class CsvHelper
    {
        public static string Serialize(List<Dictionary<string, string>> data)
        {
            if (data == null || data.Count == 0)
            {
                return string.Empty;
            }

            var headers = data.SelectMany(row => row.Keys).Distinct().ToList();
            var builder = new StringBuilder();
            builder.AppendLine(string.Join(",", headers.Select(Escape)));

            foreach (var row in data)
            {
                builder.AppendLine(string.Join(",", headers.Select(header =>
                {
                    string value;
                    row.TryGetValue(header, out value);
                    return Escape(value ?? string.Empty);
                })));
            }

            return builder.ToString();
        }

        public static List<Dictionary<string, string>> Deserialize(string content)
        {
            var result = new List<Dictionary<string, string>>();
            if (string.IsNullOrWhiteSpace(content))
            {
                return result;
            }

            var lines = content.Replace("\r\n", "\n").Split('\n').Where(line => !string.IsNullOrWhiteSpace(line)).ToList();
            if (lines.Count == 0)
            {
                return result;
            }

            var headers = ParseLine(lines[0]);
            for (var i = 1; i < lines.Count; i++)
            {
                var values = ParseLine(lines[i]);
                var row = new Dictionary<string, string>();
                for (var column = 0; column < headers.Count; column++)
                {
                    row[headers[column]] = column < values.Count ? values[column] : string.Empty;
                }

                result.Add(row);
            }

            return result;
        }

        private static List<string> ParseLine(string line)
        {
            var values = new List<string>();
            var current = new StringBuilder();
            var inQuotes = false;

            foreach (var ch in line)
            {
                if (ch == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (ch == ',' && !inQuotes)
                {
                    values.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(ch);
            }

            values.Add(current.ToString());
            return values;
        }

        private static string Escape(string value)
        {
            if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }

            return value;
        }
    }
}
