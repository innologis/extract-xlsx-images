import openpyxl
from openpyxl_image_loader import SheetImageLoader
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="XLSX Input file", required=True)
    parser.add_argument("-s", "--sheet", help="Sheet name", required=True)
    parser.add_argument("-f", "--ifieldid", help="Image field ID", type=int, required=True)
    parser.add_argument("-r", "--rfieldid", help="Relation field ID", type=int, required=True)
    parser.add_argument("-l", "--startfrom", help="Start from the line No", type=int, required=True)
    parser.add_argument("-o", "--outdir", help="Path to folder to place results in")

    args = parser.parse_args()

    od = args.outdir + "/" if args.outdir else "./"
    sl = args.startfrom
    idx = args.ifieldid
    name = args.rfieldid

    pxl_doc = openpyxl.load_workbook(args.input)
    sheet = pxl_doc[args.sheet]
    image_loader = SheetImageLoader(sheet)

    all_rows = list(sheet.rows)

    for row in all_rows[sl:]:
            print('{} {} [{}:{}]'.format(row[name].value, row[idx].value, row[name].column_letter, row[name].coordinate))
            img = image_loader.get(row[idx].coordinate)
            img.save(od + row[name].value+'.jpg')

if __name__ == "__main__":
    main()