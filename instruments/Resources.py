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
            raise SystemExit


        # Common data (usually does not need changes)
        try:
            self.blank_file = openpyxl.open("data/sample.xlsx")
        except FileNotFoundError as ex:
            print(ex)
            raise SystemExit

        # Adding colours for cosy prints
        init(autoreset=True)
        self.GREEN = lambda text: f"{Fore.GREEN}{text}{Style.RESET_ALL}"
        self.RED = lambda text: f"{Fore.RED}{text}{Style.RESET_ALL}"
        self.YELLOW = lambda text: f"{Fore.YELLOW}{text}{Style.RESET_ALL}"
        self.BLUE = lambda text: f"{Fore.BLUE}{text}{Style.RESET_ALL}"

    def close(self):
        self.export_file.close()
        self.blank_file.close()

