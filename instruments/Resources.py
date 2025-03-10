import openpyxl
from colorama import Fore, Back, Style, init


class Resources:
    def __init__(self):
        # Common data (usually does not need changes)
        try:
            self.blank_file = openpyxl.open("data/sample.xlsx")
            self.blank_sheet = self.blank_file.active
            self.book_names_data = openpyxl.open("data/names_data.xlsx")  # Empty table
            self.names_sheet = self.book_names_data.active
            self.book_empty = openpyxl.Workbook()  # Empty table
            self.empty_sheet = self.book_empty.active
        except FileNotFoundError as ex:
            print(ex)
            print("Problems with common data files load. Use init_project()\n")


        # Adding colours for cosy prints
        init(autoreset=True)
        self.GREEN = lambda text: f"{Fore.GREEN}{text}{Style.RESET_ALL}"
        self.RED = lambda text: f"{Fore.RED}{text}{Style.RESET_ALL}"
        self.YELLOW = lambda text: f"{Fore.YELLOW}{text}{Style.RESET_ALL}"
        self.BLUE = lambda text: f"{Fore.BLUE}{text}{Style.RESET_ALL}"

    def close(self):
        try:
            self.blank_file.close()
            self.book_empty.close()
            self.book_names_data.close()
        except AttributeError as ex:
            print(ex)
