import os
import time
from dotenv import load_dotenv

from instruments.Resources import Resources
from instruments.DataInstruments import DataInstruments
from instruments.parsers.feron_ua import Feron_ua


def main():
    DI.init_project()

    # Feron_ua().scrap()

    # DI.name_changer()                       # Change names according to settings
    # DI.groups_filler()                      # Autofill groups from groups_sheet (default "new_groups.xlsx")
    # DI.compress_pdf_folder()                # Compress all the files by screenshotting pages
    # DI.check_duplicates_articule("", "")    # Checks for duplicates between two files
    # DI.get_coloured_cells()                 # Checks for coloured cells in file and write them to empty file
    # DI.get_code_row()                       # Create file with codes divided by commas

    # DI.extract_unique_categories()          # Creates file with unique categories from xlsx files in folder
    # DI.download_categories(True)            # Selenium downloads categories files from categories file

if __name__ == '__main__':
    start = time.time()

    load_dotenv()
    res, DI = Resources(), DataInstruments()
    main()
    res.close()

    print(res.BLUE(f"\nTime elapsed: {time.time() - start} seconds"))