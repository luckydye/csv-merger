## CSV to XLSX converter/merger.

Merges multiple csv files into one workbook .xlsx with multiple sheets.

## Usage with Node

```bash
node . <Dir Path with CSV files> <Output Filename>.xlsx
```

## Install as CLI tool

```bash
npm run setup
```

```bash
csv2xlsx <Dir Path with CSV files> <Output Filename>.xlsx
```

## Modify the SheetNamingScheme.json

This tool uses the SheetNamingScheme.json for naming all the different Sheets.

Set "useFileNames" to true if sheets should be named after the .csv file names.
