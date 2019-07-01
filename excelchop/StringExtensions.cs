using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace excelchop
{
    public static class StringExtensions
    {
        public static string RemoveNewlines(this string text)
        {
            StringBuilder newString = new StringBuilder(text.Length);
            foreach (char t in text)
            {
                if (t != '\r' && t != '\n') newString.Append(t);
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
    }
}