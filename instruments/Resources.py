import openpyxl
from colorama import Fore, Back, Style, init


class Resources:
    def __init__(self):

        # work file load
        try:
            self.work_file = openpyxl.load_workbook("прайс_ферон.xlsx")
            self.work_sheet = self.work_file.active
        except Exception as ex:
            print(ex)
            print("Problems with work_file load\n")

        # export file load
        try:
            self.export_file = openpyxl.load_workbook("add_categories.xlsx")
            self.export_sheet = self.export_file["export_sheet"]
            self.groups_sheet = self.export_file["groups_sheet"]
        except Exception as ex:
            print(ex)
            print("Problems with add_categories.xlsx load\n")


        # Common data (usually does not need changes)
        try:
            self.blank_file = openpyxl.open("data/sample.xlsx")
            self.book_empty = openpyxl.Workbook()  # Empty table
            self.empty_sheet = self.book_empty.active
        except Exception as ex:
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
            self.export_file.close()
            self.blank_file.close()
            self.work_file.close()
        except Exception as ex:
            print(ex)
            print("Cannot close files, chech all the excel files or Resources.py\n")
