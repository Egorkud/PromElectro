import openpyxl
from colorama import Fore, Back, Style, init


class Resources:
    def __init__(self):
        try:
            self.export_file = openpyxl.load_workbook("add_categories.xlsx")
            self.export_sheet = self.export_file["export_sheet"]
            self.groups_sheet = self.export_file["groups_sheet"]
        except Exception as ex:
            print(ex)


        # Common data (usually does not need changes)
        try:
            self.blank_file = openpyxl.open("data/sample.xlsx")
        except Exception as ex:
            print(ex)

        # Adding colours for cosy prints
        init(autoreset=True)
        self.GREEN = lambda text: f"{Fore.GREEN}{text}{Style.RESET_ALL}"
        self.RED = lambda text: f"{Fore.RED}{text}{Style.RESET_ALL}"
        self.YELLOW = lambda text: f"{Fore.YELLOW}{text}{Style.RESET_ALL}"
        self.BLUE = lambda text: f"{Fore.BLUE}{text}{Style.RESET_ALL}"

    def close(self):
        try:
            self.export_file.close()
            self.blank_file.close()
        except Exception as ex:
            print(ex)
            print(self.RED("\nCannot close, check the files\n"))
