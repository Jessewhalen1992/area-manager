using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace AreaManager.Services
{
    public static class TextParsingService
    {
        private static readonly Regex NumberRegex = new Regex(@"\d+(\.\d+)?", RegexOptions.Compiled);

        public static string ExtractLastNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var matches = NumberRegex.Matches(text);
            return matches.Count > 0 ? matches[matches.Count - 1].Value : string.Empty;
        }

        public static string ExtractAllNumbers(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var matches = NumberRegex.Matches(text);
            if (matches.Count == 0)
            {
                return string.Empty;
            }

            var parts = new List<string>();
            foreach (Match match in matches)
            {
                parts.Add(match.Value);
            }

            return string.Join(" ", parts);
        }

        public static double ParseDoubleOrDefault(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            {
                return result;
            }

            return 0.0;
        }
    }
}
