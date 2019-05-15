using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace excelconvert
{
    class Program
    {
        static void Main(string[] args)
        {
            List<IOption> availableOptions = new List<IOption>()
            {
                new HelpOption(),
                new RangeOption(),
                new WorksheetOption(),
                new DelimiterOption(),
                new VersionOption(),
            };

            var argList = args.ToList();
            ConvertOptions opts = new ConvertOptions();

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                var option = availableOptions.FirstOrDefault(o => (("-" + o.ShortName) == arg || ("--" + o.LongName) == arg));
                if (option == null && i != args.Length - 1)
                {
                    Console.Out.WriteLine($"Unknown option {arg}");
                    return;
                }

                if (option == null && i == args.Length - 1)
                {
                    opts.Filename = args[i];
                    continue;
                }

                option.OptionUpdate(argList.GetRange(i, option.ArgsConsumed), opts);

                i += option.ArgsConsumed - 1;
            }

            if (opts.HelpWanted)
            {
                Console.Out.WriteLine("Usage: excelchop [options...] excel_file\n\nOptions:");

                var optionText = availableOptions.Select(option => $"    -{option.ShortName}  --{ ($"{option.LongName} {(option.ArgsConsumed == 2 ? $"<{option.LongName}>"  : "")}"),-15}    {option.HelpText}\n");
                Console.Out.Write(string.Join(string.Empty, optionText));

                Console.Out.WriteLine("\n" + HelpText.Text());
                return;
            }

            if (opts.VersionWanted)
            {
                Console.Out.WriteLine("Version 0.1");
            }

            Run(opts);
        }

        static void Run(ConvertOptions options)
        {
            string fullPath = Path.GetFullPath(options.Filename);
            if (!File.Exists(fullPath))
            {
                Console.Out.WriteLine($"Could not locate file {fullPath}.");
            }

            FileInfo fileInfo = new FileInfo(fullPath);

            using (ExcelPackage excelFile = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet sheet = options.SheetSpecified ? excelFile.Workbook.Worksheets[options.WorksheetName] : excelFile.Workbook.Worksheets.First();

                if (options.RangeSpecified)
                {
                    int numberOfColons = options.Range.Count(c => c == ':');

                    // This handles the normal cases of explicit ranges, like A1:C3 or just a
                    // single cell A2
                    if (numberOfColons < 2)
                    {
                        ExcelRange range = sheet.Cells[options.Range];

                        int startRow = range.Start.Row;
                        int endRow = range.End.Row;
                        int startColumn = range.Start.Column;
                        int endColumn = range.End.Column;

                        string output = GetOutput(options, startRow, endRow, startColumn, endColumn, sheet);
                        Console.Out.Write(output);
                    }
                    else if (numberOfColons == 2)
                    {
                        string[] rangeInputs = options.Range.Split(':');

                        bool success = int.TryParse(rangeInputs[0], out int startRow);
                        if (!success)
                        {
                            Console.Out.WriteLine($"Could not parse the start line {rangeInputs[0]} in the range specifier {options.Range}.");
                            return;
                        }

                        string startColumn = rangeInputs[1];
                        string endColumn = rangeInputs[2];

                        int startColumnInt = startColumn.ExcelColumnNameToNumber();
                        int endColumnInt = endColumn.ExcelColumnNameToNumber();
                        int columnCount = endColumnInt - startColumnInt + 1;
                        List<int> columnNumbers = Enumerable.Range(startColumnInt, columnCount).ToList();

                        Func<int, bool> rowChecker = BuildRowCheck(sheet, startColumn, endColumn);

                        int currentRow = startRow;

                        List<string> records = new List<string>();
                        while (rowChecker(currentRow))
                        {
                            int row = currentRow;
                            IEnumerable<string> fields = columnNumbers.Select(col => sheet.Cells[row, col].Text.RemoveNewlines());
                            records.Add(string.Join(options.Delimiter, fields) + "\n");
                            currentRow++;
                        }

                        string output = string.Concat(records);
                        Console.Out.Write(output);
                    }
                }
                else
                {
                    if (sheet.Dimension == null) return;

                    int startRow = sheet.Dimension.Start.Row;
                    int endRow = sheet.Dimension.End.Row;
                    int startColumn = sheet.Dimension.Start.Column;
                    int endColumn = sheet.Dimension.End.Column;

                    string output = GetOutput(options, startRow, endRow, startColumn, endColumn, sheet);
                    Console.Out.Write(output);
                }
            }
        }

        private static Func<int, string> RangeBuilder(string startColumn, string endColumn)
        {
            return row => $"{startColumn}{row}:{endColumn}{row}";
        }

        private static Func<int, bool> BuildRowCheck(ExcelWorksheet sheet, string startColumn, string endColumn)
        {
            return row => sheet.Cells[$"{startColumn}{row}:{endColumn}{row}"].Any(cell => !string.IsNullOrWhiteSpace(cell.Text));
        }


        private static string GetOutput(ConvertOptions options, int startRow, int endRow, int startColumn, int endColumn, ExcelWorksheet sheet)
        {
            List<List<string>> values = new List<List<string>>();
            for (int row = startRow; row <= endRow; row++)
            {
                values.Add(new List<string>());
                for (int column = startColumn; column <= endColumn; column++)
                {
                    // Remove all newlines as they wreck everything.
                    string cleanText = sheet.Cells[row, column].Text.RemoveNewlines();
                    values.Last().Add(cleanText);
                }
            }

            IEnumerable<string> lines = values.Select(list => string.Join(options.Delimiter, list));
            string output = string.Join("\n", lines) + "\n";
            return output;
        }


        interface IOption
        {
            char ShortName { get; }
            string LongName { get; }
            int ArgsConsumed { get; }
            void OptionUpdate(List<string> args, ConvertOptions options);
            string HelpText { get; }
        }

        public class HelpOption : IOption
        {
            public char ShortName => 'h';
            public string LongName => "help";
            public int ArgsConsumed => 1;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.HelpWanted = true;
            }

            public string HelpText => "Show help and exit";
        }

        public class RangeOption : IOption
        {
            public char ShortName => 'r';
            public string LongName => "range";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.RangeSpecified = true;
                options.Range = args.Last();
            }

            public string HelpText => "Specify range (A1:B2 or 2:A:B style)";
        }

        public class WorksheetOption : IOption
        {
            public char ShortName => 'w';
            public string LongName => "sheet";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.SheetSpecified = true;
                options.WorksheetName = args.Last();
            }

            public string HelpText => "Worksheet name";
        }

        public class DelimiterOption : IOption
        {
            public char ShortName => 'd';
            public string LongName => "delim";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.Delimiter = args.Last();
            }

            public string HelpText => "Field delimiter";
        }

        public class VersionOption : IOption
        {
            public char ShortName => 'v';
            public string LongName => "version";
            public int ArgsConsumed => 1;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.VersionWanted = true;
            }

            public string HelpText => "Show version and exit";
        }

        public class ConvertOptions
        {
            public string Filename;
            public string Range = "A1:B10";
            public bool RangeSpecified = false;
            public bool HelpWanted = false;
            public bool SheetSpecified = false;
            public string WorksheetName = "Sheet 1";
            public string Delimiter = "\t";
            public bool VersionWanted = false;
        }
    }

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

        public static int ExcelColumnNameToNumber(this string columnName)
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
    }



}
