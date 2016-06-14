# xwbp [![Build Status](https://travis-ci.org/gmontalvoriv/xwbp.svg)](https://travis-ci.org/gmontalvoriv/xwbp)

> A Node.js CLI for parsing XLSX / XLSM / XLSB / XLS / SpreadsheetML (Excel Spreadsheet) / ODS files using SheetJS's xlsx library

## Installation

```
$ npm install -g xwbp
```

## Usage

```
$ xwbp --help

  Usage: xwbp <options> <FILE>

  Options:

    -h, --help         output usage information
    -V, --version      output the version number
    --json <file>      converts a workbook object to an array of JSON objects
    --csv <file>       generates delimiter-separated-values output
    --formulae <file>  generates a list of the formulae (with value fallbacks)

  Supported read formats:

    - Excel 2007+ XML Formats (XLSX/XLSM)
    - Excel 2007+ Binary Format (XLSB)
    - Excel 2003-2004 XML Format (XML "SpreadsheetML")
    - Excel 97-2004 (XLS BIFF8)
    - Excel 5.0/95 (XLS BIFF5)
    - OpenDocument Spreadsheet (ODS)
```

Saving to a file:
```
$ xwbp --json example.xlsx > test.json
```

<i>test.json</i> :
```json
{
  "Sheet1": [
    {
      "id": "1",
      "first_name": "John",
      "last_name": "Doe",
      "stud_num": "R00123456"
    },
    {
      "id": "2",
      "first_name": "Sarah",
      "last_name": "Smith",
      "stud_num": "R00654321"
    }
  ]
}
```

## License

[MIT](https://github.com/gmontalvoriv/xwbp/blob/master/LICENSE) © Gabriel Montalvo
