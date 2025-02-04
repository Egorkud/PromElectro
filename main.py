import time

from instruments.Resources import Resources
from instruments.DataInstruments import DataInstruments
from instruments.DataScrappers import DataScrappers


def main():
    DI.init_project()

    # DI.groups_filler()                      # Autofill groups from groups_sheet (default "new_groups.xlsx")

    # DS.instructions_from_links()            # Get instructions from file with links
    # DS.photo_from_urls()                    # Get all photos from file with links
    # DS.characteristics_from_articules()     # Get all characteristics from file with articules



if __name__ == '__main__':
    start = time.time()

    res, DI, DS = Resources(), DataInstruments(), DataScrappers()
    main()
    res.close()

    print(res.BLUE(f"\nTime elapsed: {time.time() - start} seconds"))