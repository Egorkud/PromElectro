import time

from instruments.Resources import Resources
from instruments.DataInstruments import DataInstruments


def main():
    DI.init_project()

    DI.groups_filler()          # Autofill groups from groups_sheet (default "new_groups.xlsx")


if __name__ == '__main__':
    start = time.time()

    res, DI = Resources(), DataInstruments()
    main()
    res.close()

    print(res.BLUE(f"\nTime elapsed: {time.time() - start} seconds"))