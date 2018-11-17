# excelchop

`excelchop` is a command line utility to extract data out of Microsoft
Excel files.

## Motivation

I am an engineer who loves Unix utilities, but I'm forced to work in the
Microsoft environment with Excel spreadsheets being my colleagues
favorite tool. Following the Unix tradition, this program's sole job is
to get information from the Excel file to standard output.

If you get some use out of this please Star it!

# Usage

`excelchop` extracts data out of Microsoft Excel files and sends it to
standard output. From here, you can pipe the data through other filters
to achieve your goals.

By default, `excelchop` will return all the data within the first
worksheet. Using the '-r' option, you can specify a subset range. You
can either specify the range like 

`excelchop -r A1:B10 excelfile.xlsx`

or you can allow `excelchop` to automatically find the last row. You can
use the special range syntax `startrow:startcolumn:endcolumn`.

`excelchop -r 2:A:D excelfile.xlsx`

This will start at row 2, extracting data from columns A to D, stopping
once it reaches a row in which all the values are empty or whitespace.

The default delimiter is a tab character and output records are
separated by a Unix newline. `excelchop` also removes any newline
characters within a field.

# Installation

See the release pages for downloads.

## Windows

1. Extract the `excelchop-windows.zip` file.
2. Add the extracted folder to your `PATH` environment variable.

## Unix

1. Extract the `excelchop-unix.zip` file.
2. Symlink the `excelchop` binary to a location in `PATH`.

