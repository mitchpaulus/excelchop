# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build Commands

```bash
# Build the solution
dotnet build excelchop.sln

# Build release
dotnet build excelchop.sln -c Release

# Run tests
dotnet test tests/tests.csproj

# Run a single test
dotnet test tests/tests.csproj --filter "FullyQualifiedName~TestSigFigs"

# Run the CLI tool
dotnet run --project excelchop/excelchop.csproj -- [args]
```

## Architecture

This is a .NET 9 command-line tool that extracts data from Excel files to stdout using the EPPlus library.

**Project structure:**
- `excelchop/` - Main CLI application
- `tests/` - NUnit test project

**Key files:**
- `Program.cs` - Entry point with command-line parsing and main extraction logic. Contains all option classes implementing `IOption` interface for argument handling.
- `ExcelUtilities.cs` - Cell reference parsing (A1 and R1C1 notation) with `TryParseCellReference`
- `StringExtensions.cs` - Helper methods for column name conversion (`ExcelColumnNameToInt`), newline handling, and output formatting

**Command-line option pattern:**
Options implement `IOption` interface with `ShortName`, `LongName`, `ArgsConsumed`, and `OptionUpdate` method. Add new options to the `availableOptions` list in `Main()`.

**Range formats supported:**
- Single cell: `A1`
- Explicit range: `A1:B10`
- Auto-find last row: `2:A:D` (row:startcol:endcol) - stops on blank rows based on `-A`, `-s`, `-S` options
