# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- New `get_shared_string` function to resolve indexes of shared strings.

## [0.1.1]

### Fixed

- Prevented XLSX cell corruption when writing numeric values into cells that
  previously had string/shared-string type (`t` attribute). Numeric writes now
  remove `t` so the written `<v>` is interpreted correctly by spreadsheet apps.

## [0.1.0]

Initial release

Available features:

- Parsing
- Single Cell reading
- Single cell writing
- Exporting

[Unreleased]: https://github.com/stritzinger/xlerl/compare/0.1.1...HEAD
[0.1.1]: https://github.com/stritzinger/xlerl/compare/0.1.0...0.1.1
[0.1.0]: https://github.com/stritzinger/xlerl/compare/ec3038933b203555d30ac4a0d948f2b862718433...0.1.0
