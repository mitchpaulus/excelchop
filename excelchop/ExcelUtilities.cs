using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace excelchop
{
    public static class ExcelUtilities
    {
        public static bool TryParseCellReference(string cellReference, [NotNullWhen(true)] out Cell? cellLocation)
        {
            string worksheetNamePattern = @"'[^:\\/?*[\]]{1,31}'!";
            string worksheetNumberPattern = @"[1-9]\d*!";
            Regex a1Regex = new Regex($@"^(({worksheetNumberPattern})|({worksheetNamePattern}))?([A-Za-z]+)([0-9]+)$");
            Regex r1c1Regex = new Regex("^[rR]([0-9]+)[cC]([0-9]+)$");

            var a1Match = a1Regex.Match(cellReference);
            var r1c1Match = r1c1Regex.Match(cellReference);

            if (a1Match.Success)
            {
                cellLocation = new Cell (
                    sheetName: a1Match.Groups[3].Success ? a1Match.Groups[3].Value.Substring(1, a1Match.Groups[3].Value.Length - 3) : null,
                    sheetNum:  a1Match.Groups[2].Success ? int.Parse(a1Match.Groups[2].Value.Substring(0, a1Match.Groups[2].Value.Length - 1)) : -1,
                    row:  int.Parse(a1Match.Groups[5].Value),
                    column: a1Match.Groups[4].Value.ExcelColumnNameToInt()
                );
                return true;
            }
            else if (r1c1Match.Success)
            {
                cellLocation = new Cell
                (
                    null,
                    -1,
                    int.Parse(r1c1Match.Groups[1].Value),
                    int.Parse(r1c1Match.Groups[2].Value)
                );
                return true;
            }
            else
            {
                cellLocation = null;
                return false;
            }
        }
    }

    public class Cell
    {
        public string? SheetName;
        public int SheetNum;
        public int Row;
        public int Column;

        public Cell(string? sheetName, int sheetNum, int row, int column)
        {
            SheetName = sheetName;
            SheetNum = sheetNum;
            Row = row;
            Column = column;
        }
    }

}
