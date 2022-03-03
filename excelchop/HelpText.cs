using System.Text;

namespace excelchop
{
    public static class HelpText
    {
        public static string Text()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("excelchop extracts data out of Microsoft Excel files and sends it to\n");
            sb.Append("standard output. From here, you can pipe the data through other filters\n");
            sb.Append("to achieve your goals.\n");
            sb.Append("\n");
            sb.Append("By default, excelchop will return all the data within the first\n");
            sb.Append("worksheet. Using the '-r' option, you can specify a subset range. You\n");
            sb.Append("can either specify the range like\n");
            sb.Append("\n");
            sb.Append("excelchop -r A1:B10 excelfile.xlsx\n");
            sb.Append("\n");
            sb.Append("or you can allow excelchop to automatically find the last row. You can\n");
            sb.Append("use the special range syntax 'startrow:startcolumn:endcolumn'.\n");
            sb.Append("\n");
            sb.Append("excelchop -r 2:A:D excelfile.xlsx\n");
            sb.Append("\n");
            sb.Append("This will start at row 2, extracting data from columns A to D, stopping\n");
            sb.Append("once it reaches a row in which ANY of the values are empty or\n");
            sb.Append("whitespace. You can use the options -A, -s, or -S, to change this\n");
            sb.Append("stopping behavior.\n");
            sb.Append("\n");
            sb.Append("The default delimiter is a tab character and output records are\n");
            sb.Append("separated by a Unix newline. excelchop also removes any newline\n");
            sb.Append("characters within a field.\n");

            return sb.ToString();
        }
    }
}
