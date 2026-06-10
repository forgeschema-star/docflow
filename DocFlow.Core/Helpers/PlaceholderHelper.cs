using System;
using System.Collections.Generic;

namespace DocFlow.Core.Helpers
{
    public static class PlaceholderHelper
    {
        public static void ValidatePlaceholders(IDictionary<string, string> placeholders, string parameterName)
        {
            if (placeholders == null)
            {
                throw new ArgumentNullException(parameterName);
            }
        }

        public static string ReplacePlaceholders(string input, IDictionary<string, string> placeholders)
        {
            if (input == null)
            {
                return string.Empty;
            }

            ValidatePlaceholders(placeholders, nameof(placeholders));

            var output = input;
            foreach (var item in placeholders)
            {
                var key = item.Key ?? string.Empty;
                var value = item.Value ?? string.Empty;

                output = output.Replace("{{" + key + "}}", value);
                output = output.Replace("[[" + key + "]]", value);
                output = output.Replace("<<" + key + ">>", value);
                output = output.Replace("{" + key + "}", value);
            }

            return output;
        }
    }
}
