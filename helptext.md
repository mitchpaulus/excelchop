excelchop extracts data out of Microsoft Excel files and sends it to
standard output. From here, you can pipe the data through other filters
to achieve your goals.

By default, excelchop will return all the data within the first
worksheet. Using the '-r' option, you can specify a subset range. You
can either specify the range like 

excelchop -r A1:B10 excelfile.xlsx

or you can allow excelchop to automatically find the last row. You can
use the special range syntax 'startrow:startcolumn:endcolumn'.

excelchop -r 2:A:D excelfile.xlsx

This will start at row 2, extracting data from columns A to D, stopping
once it reaches a row in which all the values are empty or whitespace.

The default delimiter is a tab character and output records are
separated by a unix newline. excelchop also removes any newline
characters within a field.

