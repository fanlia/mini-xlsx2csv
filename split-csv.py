
import csv, sys
import argparse

def split_csv(options):
    print(options)
    xlsxfile = options.xlsxfile
    if xlsxfile == '-':
        csvfile = sys.stdin
    else:
        csvfile = open(xlsxfile)

    reader = csv.DictReader(csvfile)
    writers = {}
    maxfiles = 0
    for row in reader:
        value = row.get(options.field)
        if not value:
            print('field not found')
            break

        writer = writers.get(value)
        if not writer:
            if maxfiles >= options.limit:
                continue
            keys = row.keys()
            filename = value + '.csv'
            file = open(filename, 'w')
            print('created file', filename)
            csvwriter = csv.DictWriter(file, keys)
            csvwriter.writerow({key:key for key in keys})
            writer = (file, csvwriter)
            writers[value] = writer
            maxfiles += 1

        writer[1].writerow(row)

    for (file, writer) in writers.values():
        file.close()

    csvfile.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='split csv by field')
    parser.add_argument('xlsxfile', help='xlsx file path, - means from stdin')
    parser.add_argument('field', help='field to split file')
    parser.add_argument('--limit', dest='limit', help='max files to write', type=int, default=10)
    options = parser.parse_args()

    split_csv(options)
