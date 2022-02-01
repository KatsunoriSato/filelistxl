import openpyxl
import os
import sys
import glob

def createFileList():
    xlrow = 1
    xlcolumn = 1

    # create book file
    book = openpyxl.Workbook()

    # get sheet
    sheet = book['Sheet']
    
    # target path
    path = sys.argv[1]

    sheet.cell(xlrow, xlcolumn).value = "Directory:" + path
    xlrow += 1

    files = glob.glob(path + "/*")
    for file in files:
        sheet.cell(xlrow, xlcolumn).value = file
        xlrow += 1

    # rename sheet
    sheet.title = 'filelist'

    # save file
    book.save("filelist.xlsx")


# main flow
def main():
    argc = len(sys.argv)
    if argc == 1:
        print('-- filelistxl --')
        print('>python3 filelistxl [path]')
        sys.exit(0)
    
    if os.path.exists(sys.argv[1]) == False:
        print('filelistxl ')
        print('>python3 filelistxl [path]')
        sys.exit(0)

    createFileList()



if __name__ == "__main__":
    main()

