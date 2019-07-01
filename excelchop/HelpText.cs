using System.Text;

namespace excelchop
{
    public static class HelpText
    {
        public static string Text()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("excelchop extracts data out of Microsoft Excel files and sends it to");
            sb.AppendLine("standard output. From here, you can pipe the data through other filters");
            sb.AppendLine("to achieve your goals.");
            sb.AppendLine("");
            sb.AppendLine("By default, excelchop will return all the data within the first");
            sb.AppendLine("worksheet. Using the '-r' option, you can specify a subset range. You");
            sb.AppendLine("can either specify the range like");
            sb.AppendLine("");
            sb.AppendLine("excelchop -r A1:B10 excelfile.xlsx");
            sb.AppendLine("");
            sb.AppendLine("or you can allow excelchop to automatically find the last row. You can");
            sb.AppendLine("use the special range syntax 'startrow:startcolumn:endcolumn'.");
            sb.AppendLine("");
            sb.AppendLine("excelchop -r 2:A:D excelfile.xlsx");
            sb.AppendLine("");
            sb.AppendLine("This will start at row 2, extracting data from columns A to D, stopping");
            sb.AppendLine("once it reaches a row in which ANY of the values are empty or");
            sb.AppendLine("whitespace. You can use the options -A, -s, or -S, to change this");
            sb.AppendLine("stopping behavior.");
            sb.AppendLine("");
            sb.AppendLine("The default delimiter is a tab character and output records are");
            sb.AppendLine("separated by a Unix newline. excelchop also removes any newline");
            sb.AppendLine("characters within a field.");
            sb.AppendLine("");


            return sb.ToString();
        }
    }
}
