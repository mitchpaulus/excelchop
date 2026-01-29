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
characters within a field. Output is UTF-8 encoded without a BOM.

# Installation

See the [release page](https://github.com/mitchpaulus/excelchop/releases) for downloads.
There are releases for Windows, Linux, and macOS.
For each operating system there is a standalone and framework dependent build.
The self contained versions will run without any other prerequisites, but are larger in size.
The framework-dependent versions require the .NET runtime be already installed.

1. Extract the zip file.
2. Move or symlink the executable file to a location in your `PATH` environment variable.

## Related Projects

- [`xlsx2csv`](https://github.com/tealeg/xlsx2csv)
