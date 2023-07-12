# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [0.8.0](https://github.com/mitchpaulus/excelchop/compare/v0.7.0...v0.8.0) - 2023-07-12

### Breaking

- By default, string content is stripped of leading and trailing whitespace.
  This can be disabled with the new option `--no-strip`.

### Fixed

- Error in logic when dealing with single cell range like `A1`.

## [0.7.0](https://github.com/mitchpaulus/excelchop/compare/v0.6.0...v0.7.0) - 2023-05-25

### Added

- New option `--in-place` or `-i` to write the data to a new TSV file with the same name as the input Excel file.

## [0.6.0](https://github.com/mitchpaulus/excelchop/compare/v0.5.0...v0.6.0) - 2022-11-29

### Changed

- Rebuilt targeting .NET 7.

## [0.5.0] - 2022-11-02

### Added

- New option `--all-sheets`. If you want data from every sheet, you no longer need to specify each sheet directly in the `-w` argument.

## [0.4.0] - 2022-10-28

### Added

- The worksheet option `-w` now allows for multiple worksheets to be entered.
  The argument that follows the `-w` is now split on ':' or newlines.

### Changed

- The structure of the releases has been updated for the better.
  Assets for the releases are now built using GitHub actions.

## [0.3.0] - 2022-05-20

### Added

- New option for significant figures.
  Can help with downstream programs that cannot handle thousands separators.


[0.3.0]: https://github.com/mitchpaulus/excelchop/compare/v0.2.3...v0.3.0
[0.4.0]: https://github.com/mitchpaulus/excelchop/compare/v0.3.0...v0.4.0
[0.5.0]: https://github.com/mitchpaulus/excelchop/compare/v0.4.0...v0.5.0
