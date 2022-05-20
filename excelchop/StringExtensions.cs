using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace excelchop
{
    public static class StringExtensions
    {
        public static string EscapeNewlines(this string text)
        {
            StringBuilder newString = new StringBuilder(text.Length);
            foreach (char t in text)
            {
                // Remove carriage returns
                if (t == '\r') continue;

                if (t == '\n') newString.Append("\\n");
                else newString.Append(t);
            }
            return newString.ToString();
        }

        public static int ExcelColumnNameToInt(this string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException(nameof(columnName));

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            foreach (char c in columnName)
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }

            return sum;
        }

        public static List<int> ColsFromList(this string optionList)
        {
            return optionList.Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(s => int.TryParse(s.Trim(), out int col) ? col : s.Trim().ExcelColumnNameToInt())
            .ToList();
        }

        public static string EndWithSingleNewline(this string? inputString)
        {
            // Handle null or empty strings
            if (string.IsNullOrEmpty(inputString)) return "";

            // If a string also has multiple newlines, return with just 1 at the end
            var i = inputString.Length - 1;

            while (i >= 0 && inputString[i] == '\n') i--;
            return inputString.Substring(0, i + 1) + "\n";
        }
    }
}
