# mini-xlsx2csv

just a mini verson of the great work from https://github.com/dilshod/xlsx2csv

## usage

### xlsx to csv

```sh
usage: mini-xlsx2csv.py [-h] [--limit LIMIT] [--sheetname SHEETNAME] [--field FIELD] xlsxfile

xlsx as csv to stdout

positional arguments:
  xlsxfile              xlsx file path

optional arguments:
  -h, --help            show this help message and exit
  --limit LIMIT         rows to write
  --sheetname SHEETNAME
                        sheet name to convert
  --field FIELD         field to extrac
```

### split csv

```sh
usage: split-csv.py [-h] [--limit LIMIT] xlsxfile field

split csv by field

positional arguments:
  xlsxfile       xlsx file path, - means from stdin
  field          field to split file

optional arguments:
  -h, --help     show this help message and exit
  --limit LIMIT  max files to write

```
## why

- only support python3
- fixed datetime format: %Y-%m-%d %H:%M:%S
- csv goes to stdout

## thanks

https://github.com/dilshod/xlsx2csv

## license

GPL

