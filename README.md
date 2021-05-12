# mini-xlsx2csv

just a mini verson of the great work from https://github.com/dilshod/xlsx2csv

## usage

```sh
usage: mini-xlsx2csv.py [-h] [--limit LIMIT] [--sheetname SHEETNAME] xlsxfile

xlsx as csv to stdout

positional arguments:
  xlsxfile              xlsx file path

optional arguments:
  -h, --help            show this help message and exit
  --limit LIMIT         rows to write
  --sheetname SHEETNAME
                        sheet name to convert
```

## why

- only support python3
- fixed datetime format: %Y-%m-%d %H:%M:%S
- csv goes to stdout

## thanks

https://github.com/dilshod/xlsx2csv

## license

GPL

