import time

from instruments.Resources import Resources
from instruments.DataInstruments import DataInstruments


def main():
    ...



if __name__ == '__main__':
    start = time.time()

    res, DI = Resources(), DataInstruments()
    main()
    res.close()

    print(res.BLUE(f"\nTime elapsed: {time.time() - start} seconds"))