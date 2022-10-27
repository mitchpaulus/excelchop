using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace excelchop
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
                new DateTimeFormatOption(),
                new EscapeNewLinesOption(),
                new VersionOption(),
                new AllFieldsAllBlankOption(),
                new StopAnyOption(),
                new StopAllOption(),
                new PrintInfoOption(),
                new SigFigsOption(),
            };

            var argList = args.ToList();
            ConvertOptions opts = new ConvertOptions();

            bool OptionMatches(string arg, IOption option)
            {
                if (option.ShortName != null && "-" + option.ShortName == arg) return true;
                return "--" + option.LongName == arg;
            }

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                var option = availableOptions.FirstOrDefault(o => OptionMatches(arg, o));
                if (option == null && i != args.Length - 1)
                {
                    Console.Error.Write($"Unknown option {arg}\n");
                    Environment.ExitCode = 1;
                    return;
                }

                if (option == null && i == args.Length - 1)
                {
                    opts.Filename = args[i];
                    continue;
                }

                try
                {
                    option!.OptionUpdate(argList.GetRange(i, option.ArgsConsumed), opts);
                }
                catch (ExcelChopKnownException ex)
                {
                    Environment.ExitCode = 1;
                    Console.Error.Write(ex.Message);
                }

                i += option!.ArgsConsumed - 1;
            }

            if (opts.HelpWanted)
            {
                Console.Out.Write("Usage: excelchop [options...] excel_file\n\nOptions:\n");

                IEnumerable<string> optionText = availableOptions.OrderBy(option => option.ShortName).Select(FormatOptionHelpLine);
                Console.Out.Write(string.Join(string.Empty, optionText));
                Console.Out.Write("\n" + HelpText.Text());
                return;
            }

            if (opts.VersionWanted)
            {
                Console.Out.Write("0.4.0 - 2022-10-27\n");
                return;
            }

            try
            {
                Run(opts);
            }
            catch (ExcelChopKnownException knownException)
            {
                Environment.ExitCode = 1;
                Console.Error.Write(knownException.Message.EndWithSingleNewline());
            }
            catch
            {
                Environment.ExitCode = 1;
                Console.Error.Write("There was an unhandled exception. Please feel free to open an issue on GitHub at https://github.com/mitchpaulus/excelchop.\n");
            }
        }

        private static void Run(ConvertOptions options)
        {
            if (options.Filename == null)
            {
                Console.Error.Write("A file path to an existing Excel file needs to be provided.\n");
                Environment.ExitCode = 1;
                return;
            }

            string fullPath = Path.GetFullPath(options.Filename);
            if (!File.Exists(fullPath))
            {
                Console.Error.Write($"Could not locate file '{fullPath}'.\n");
                Environment.ExitCode = 1;
                return;
            }

            FileInfo fileInfo = new FileInfo(fullPath);

            List<ExcelWorksheet> sheets = new List<ExcelWorksheet>();

            using (ExcelPackage excelFile = new ExcelPackage(fileInfo))
            {

                if (options.PrintOption == PrintOption.Worksheets)
                {
                    Console.Out.Write(string.Concat(excelFile.Workbook.Worksheets.Select(s => s.Name + '\n')));
                    return;
                }

                //ExcelWorksheet sheet;
                if (options.SheetSpecified)
                {
                    HashSet<string> availableWorksheetNames = excelFile.Workbook.Worksheets.Select(worksheet => worksheet.Name).ToHashSet();

                    foreach (string worksheetName in options.WorksheetNames)
                    {
                        if (!availableWorksheetNames.Contains(worksheetName))
                        {
                            Console.Error.Write($"No worksheet named {worksheetName} found in {fullPath}.\n");
                            Environment.ExitCode = 1;
                            return;
                        }

                        sheets.Add(excelFile.Workbook.Worksheets[worksheetName]);
                    }
                }
                else
                {
                    if (excelFile.Workbook.Worksheets.Any()) sheets = new List<ExcelWorksheet> {excelFile.Workbook.Worksheets.First()};
                    else
                    {
                        Console.Error.Write($"There are no worksheets in {fullPath}.\n");
                        Environment.ExitCode = 1;
                        return;
                    }
                }

                foreach (ExcelWorksheet sheet in sheets)
                {
                    if (options.RangeSpecified)
                    {
                        var splitRange = options.Range.Split(':');

                        // This handles the normal cases of explicit ranges, like A1:C3 or just a
                        // single cell A2
                        if (splitRange.Length == 1)
                        {
                            var success = ExcelUtilities.TryParseCellReference(options.Range, out Cell? cellLocation);

                            if (success)
                            {
                                string output = GetOutput(options, cellLocation!.Row, cellLocation.Column, cellLocation.Row, cellLocation.Column, sheet);
                                Console.Out.Write(output);
                            }
                            else
                            {
                                Console.Error.Write($"Could not parse cell reference {options.Range}.\n");
                                Environment.ExitCode = 1;
                            }
                        }
                        else if (splitRange.Length == 2)
                        {
                            var firstCellSuccess = ExcelUtilities.TryParseCellReference(splitRange[0], out Cell? startCellLocation);
                            var secondCellSuccess = ExcelUtilities.TryParseCellReference(splitRange[1], out Cell? endCellLocation);

                            if (firstCellSuccess && secondCellSuccess)
                            {
                                string output = GetOutput(options, startCellLocation!.Row, endCellLocation!.Row, startCellLocation.Column, endCellLocation.Column, sheet);
                                Console.Out.Write(output);
                            }
                            else
                            {
                                Console.Error.Write($"Could not parse cell reference {options.Range}.\n");
                                Environment.ExitCode = 1;
                            }
                        }
                        else if (splitRange.Length == 3)
                        {
                            bool success = int.TryParse(splitRange[0], out int startRow);
                            if (!success)
                            {
                                Console.Error.Write($"Could not parse the start line {splitRange[0]} in the range specifier {options.Range}.\n");
                                Environment.ExitCode = 1;
                                return;
                            }

                            string startColumn = splitRange[1];
                            string endColumn = splitRange[2];

                            int startColumnInt = startColumn.ExcelColumnNameToInt();
                            int endColumnInt = endColumn.ExcelColumnNameToInt();
                            int columnCount = endColumnInt - startColumnInt + 1;
                            List<int> columnNumbers = Enumerable.Range(startColumnInt, columnCount).ToList();

                            Func<int, bool> rowInvalid = options.RowInvalid(sheet, startColumn, endColumn);

                            int currentRow = startRow;

                            List<string> records = new List<string>();
                            while (!rowInvalid(currentRow))
                            {
                                int row = currentRow;
                                IEnumerable<string> fields;
                                if (options.EscapeNewlines)
                                {
                                    fields = columnNumbers.Select(col => CellText(sheet.Cells[row, col], options.DateFormat, options.SignificantDigits).EscapeNewlines());
                                }
                                else
                                {
                                    fields = columnNumbers.Select(col =>
                                        CellText(sheet.Cells[row, col], options.DateFormat, options.SignificantDigits)
                                            .Replace("\r", "")
                                            .Replace("\n", " "));
                                }
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
        }

        public static string FormatOptionHelpLine(IOption option)
        {
            return option.ShortName == null ?
                $"        --{($"{option.LongName} {(option.ArgsConsumed == 2 ? $"<{option.LongName}>" : "")}"),-23}   {option.HelpText}\n" :
                $"    -{option.ShortName}  --{($"{option.LongName} {(option.ArgsConsumed == 2 ? $"<{option.LongName}>" : "")}"),-23}   {option.HelpText}\n";
        }

        private static Func<int, string> RangeBuilder(string startColumn, string endColumn)
        {
            return row =>  $"{startColumn}{row}:{endColumn}{row}";
        }

        private static Func<int, bool> AllFieldsAnyBlank(ExcelWorksheet sheet, string startColumn, string endColumn)
        {
            return row =>
            {
                int startCol = startColumn.ExcelColumnNameToInt();
                int endCol = endColumn.ExcelColumnNameToInt();

                for (int col  = startCol;  col <= endCol; col++)
                {
                    if (string.IsNullOrWhiteSpace(sheet.Cells[row, col].Text)) return true;
                }

                return false;
            };
        }
        private static Func<int, bool> AllFieldsAllBlank(ExcelWorksheet sheet, string startColumn, string endColumn)
        {
            return row =>
            {
                int startCol = startColumn.ExcelColumnNameToInt();
                int endCol = endColumn.ExcelColumnNameToInt();

                for (int col  = startCol;  col <= endCol; col++)
                {
                    if (!string.IsNullOrWhiteSpace(sheet.Cells[row, col].Text)) return false;
                }

                return true;
            };
        }

        private static string CellText(ExcelRangeBase range, string dateFormat, int? significantDigits)
        {
            if (range.Value is DateTime dateCell)
            {
                try
                {
                    string dateText = dateCell.ToString(dateFormat);
                    return dateText;
                }
                catch (FormatException)
                {
                    throw new ExcelChopKnownException($"The datetime format '{dateFormat}' was not valid.");
                }
            }

            if (significantDigits != null && range.Value is double doubleCell)
            {
                return doubleCell.ToSigFigs((int) significantDigits);
            }

            return range.Text;
        }

        private static string GetOutput(ConvertOptions options, int startRow, int endRowInc, int startColumn, int endColumnInc, ExcelWorksheet sheet)
        {
            List<List<string>> values = new List<List<string>>();
            for (int row = startRow; row <= endRowInc; row++)
            {
                values.Add(new List<string>());
                for (int column = startColumn; column <= endColumnInc; column++)
                {
                    string cleanText;
                    // Remove all newlines as they wreck everything.
                    if (options.EscapeNewlines)
                    {
                        cleanText = CellText(sheet.Cells[row, column], options.DateFormat, options.SignificantDigits).EscapeNewlines();
                    }
                    else
                    {
                        cleanText = CellText(sheet.Cells[row, column], options.DateFormat, options.SignificantDigits)
                            .Replace("\r", "")
                            .Replace("\n", " ");
                    }
                    values.Last().Add(cleanText);
                }
            }

            IEnumerable<string> lines = values.Select(list => string.Join(options.Delimiter, list) + "\n");
            string output = string.Concat(lines);
            return output;
        }

        public interface IOption
        {
            char? ShortName { get; }
            string LongName { get; }
            int ArgsConsumed { get; }
            void OptionUpdate(List<string> args, ConvertOptions options);
            string HelpText { get; }
        }

        public class HelpOption : IOption
        {
            public char? ShortName => 'h';
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
            public char? ShortName => 'r';
            public string LongName => "range";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.RangeSpecified = true;
                options.Range = args.Last();
            }

            public string HelpText => $"Specify range (A1:B2 or 2:A:B style) [{new ConvertOptions().Range}]";
        }

        public class WorksheetOption : IOption
        {
            public char? ShortName => 'w';
            public string LongName => "sheet";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.SheetSpecified = true;
                // ':' is a character that cannot occur in a Worksheet name, so acts as a useful separator.
                options.WorksheetNames = args.Last().Split(":").ToList();
            }

            public string HelpText => "Worksheet name [first sheet]";
        }

        public class DelimiterOption : IOption
        {
            public char? ShortName => 'd';
            public string LongName => "delim";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.Delimiter = args.Last();
            }

            public string HelpText => "Output field delimiter [tab]";
        }

        public class DateTimeFormatOption : IOption
        {
            public char? ShortName => 'D';

            public string LongName => "dateformat";

            public int ArgsConsumed => 2;

            public void OptionUpdate(List<string> args, ConvertOptions options) => options.DateFormat = args.Last();

            public string HelpText => "Output format for date cells, .NET style [yyyy-MM-dd]";
        }

        public class SigFigsOption : IOption
        {
            public char? ShortName => null;
            public string LongName => "sigfigs";
            public int ArgsConsumed => 2;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                try
                {
                    options.SignificantDigits = int.Parse(args.Last());
                }
                catch
                {
                    throw new ExcelChopKnownException($"Could not parse '{args.Last()}' as an integer.");
                }
            }

            public string HelpText => "Integer significant figures for numeric cells. Default is to print using current excel format.";
        }

        public class VersionOption : IOption
        {
            public char? ShortName => 'v';
            public string LongName => "version";
            public int ArgsConsumed => 1;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.VersionWanted = true;
            }

            public string HelpText => "Show version and exit";
        }

        public class AllFieldsAllBlankOption : IOption
        {
            public char? ShortName => 'A';

            public string LongName => "all-fields-all-blank";

            public int ArgsConsumed => 1;

            public void OptionUpdate(List<string> args, ConvertOptions options) => options.RowInvalid = AllFieldsAllBlank;

            public string HelpText => "Stop reading when all fields in complete range are blank before stopping.";
        }

        public class StopAllOption : IOption
        {
            public char? ShortName => 's';

            public string LongName => "stop-all";

            public int ArgsConsumed => 2;

            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                List<int> columns = args.Last().ColsFromList();

                options.RowInvalid = (sheet, start, end) => { return row => columns.All(col => string.IsNullOrWhiteSpace(sheet.Cells[row, col].Text)); };
            }

            public string HelpText => "Stop reading when all columns specified are empty. Specify columns as comma separated list.";
        }

        public class StopAnyOption : IOption
        {
            public char? ShortName => 'S';

            public string LongName => "stop-any";

            public int ArgsConsumed => 2;

            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                List<int> columns = args.Last().ColsFromList();

                options.RowInvalid = (sheet, start, end) => { return row => columns.Any(col => string.IsNullOrWhiteSpace(sheet.Cells[row, col].Text)); };
            }

            public string HelpText => "Stop reading when any columns specified are empty. Specify columns as comma separated list.";
        }

        public class PrintInfoOption : IOption
        {
            public char? ShortName => 'p';
            public string LongName => "print";
            public int ArgsConsumed => 2;
            public string HelpText => "Print information about workbook. w = worksheet names";
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.PrintOption = PrintOption.Worksheets;
            }
        }

        public class EscapeNewLinesOption : IOption
        {
            public char? ShortName => 'e';
            public string LongName => "escape";
            public int ArgsConsumed => 1;
            public void OptionUpdate(List<string> args, ConvertOptions options)
            {
                options.EscapeNewlines = true;
            }

            public string HelpText => "Escape newlines with '\\n' characters";
        }

        public enum PrintOption
        {
            Data = 0,
            Worksheets = 1
        }

        public class ConvertOptions
        {
            public string? Filename;
            public string Range = "A1:B10";
            public bool RangeSpecified = false;
            public bool HelpWanted = false;
            public bool SheetSpecified = false;
            public List<string> WorksheetNames = new();
            public string Delimiter = "\t";
            public bool VersionWanted = false;
            public string DateFormat = "yyyy-MM-dd";
            public int? SignificantDigits = null;
            public RowCheckFunction RowInvalid = AllFieldsAnyBlank;
            public PrintOption PrintOption = PrintOption.Data;
            public bool EscapeNewlines = false;
        }

        public delegate Func<int, bool> RowCheckFunction(ExcelWorksheet sheet, string startColumnInc, string endColumnInc);

    }

    public class ExcelChopKnownException : Exception
    {
        private readonly string _message;

        public ExcelChopKnownException(string message)
        {
            _message = message;
        }
        public override string Message => _message;
    }

    public static class NumericExtensions
    {
        public static string ToSigFigs(this double value, int sigFigs)
        {
            if (value == 0) return "0";
            int log = (int)Math.Floor(Math.Log10(Math.Abs(value)));
            var digits = Math.Max(0, sigFigs - 1 - log);

            string fullString = value.ToString("f" + digits);

            if (digits == 0) return fullString;

            // Trim trailing zeros when there is a decimal
            int i = fullString.Length - 1;
            while (i > 0 && fullString[i] == '0')
            {
                i--;
            }
            if (fullString[i] == '.')
            {
                i--;
            }
            return fullString.Substring(0, i + 1);
        }
    }
}
