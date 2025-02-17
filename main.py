import time

from instruments.Resources import Resources
from instruments.DataInstruments import DataInstruments


def main():
    DI.init_project()

    # DI.name_changer()

    # DI.groups_filler()                      # Autofill groups from groups_sheet (default "new_groups.xlsx")
    # DI.compress_pdf_folder()                # Compress all the files by screenshotting pages
    # DI.check_duplicates_articule("", "")    # Checks for duplicates between two files
    # DI.get_coloured_cells()                 # Checks for coloured cells in file and write them to empty file


if __name__ == '__main__':
    start = time.time()

    res, DI = Resources(), DataInstruments()
    main()
    res.close()

    print(res.BLUE(f"\nTime elapsed: {time.time() - start} seconds"))