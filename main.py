import time
from dotenv import load_dotenv

from instruments.Resources import Resources
from instruments.DataInstruments import DataInstruments


def main():
    DI.init_project()

    # DI.name_changer()                       # Change names according to settings
    # DI.groups_filler()                      # Autofill groups from groups_sheet (default "new_groups.xlsx")
    # DI.compress_pdf_folder()                # Compress all the files by screenshotting pages
    # DI.check_duplicates_articule("", "")    # Checks for duplicates between two files
    # DI.get_coloured_cells()                 # Checks for coloured cells in file and write them to empty file
    # DI.get_code_row()                       # Create file with codes divided by commas
    # DI.generate_numbers_string(1, 5)        # Creates .txt file with numbers divided by ,
    # DI.collect_product_numbers()            # Creates .txt file with numbers divided by , from directory

    # DI.merge_xlsx_files()                   # Merges files after parsing into similar columns

    # DI.extract_unique_categories()          # Creates file with unique categories from xlsx files in folder
    # DI.download_categories(True)            # Selenium downloads categories files from categories file
    # DI.process_excel_files(directory="", )  # Process all files from dir

if __name__ == '__main__':
    start = time.time()

    load_dotenv()
    res, DI = Resources(), DataInstruments()
    main()
    res.close()

    print(res.BLUE(f"\nTime elapsed: {time.time() - start} seconds"))