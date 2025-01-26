import random
import time
import requests
from bs4 import BeautifulSoup
import openpyxl

from instruments.Resources import Resources


class DataInstruments(Resources):
    def __init__(self):
        super().__init__()

    def init_project(self):
        ...

    # Fill descriptions from descriptions sheet.
    # Column 1. Name or id as convenient
    # Column 2. Group name (full path to group).
    def groups_filler(self, filename : str = "new_groups.xlsx", export_file : str = "export.xlsx"):
        groups_dict = {}
        export_file = openpyxl.open(export_file)
        export_sheet = export_file["export sheet"]
        groups_sheet = export_file["groups sheet"]

        for row in range(1, groups_sheet.max_row + 1):
            id_name = groups_sheet.cell(row, 1).value
            group_name = groups_sheet.cell(row, 2).value
            groups_dict.update([(id_name, group_name)])

        for row in range(2, export_sheet.max_row + 1):
            id_name = export_sheet.cell(row, 3).value
            if id_name in groups_dict.keys():
                group_name = groups_dict[id_name]
                export_sheet.cell(row, 3).value = group_name
                print(self.GREEN(f"{row}. changed"))
            else:
                print(self.YELLOW(f"{row}. skipped"))

        export_file.save(filename)
        print(self.GREEN(f"\nFile {filename} created"))