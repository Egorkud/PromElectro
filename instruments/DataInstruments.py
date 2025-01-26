import random
import time
import requests
from bs4 import BeautifulSoup

from instruments.Resources import Resources


class DataInstruments(Resources):
    def __init__(self):
        super().__init__()

    def data_scrapper(self):
        for row in range(2, self.export_sheet.max_row + 1):
            print(row)
            URL = self.export_sheet.cell(row, 2).value
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
                "Cookie": "0934b93dc6fa38f78603e0d9b24017ff=ed2e533ffdfe32cc1848eeb36e0d33da",
            }

            try:
                req = requests.get(URL, headers=headers)
                src = req.text
                soup = BeautifulSoup(src, "lxml")

                # Implement data scrapping here

            except Exception as ex:
                print(ex)

            self.export_file.save("scrapped_data.xlsx")
            time.sleep(10 + random.uniform(1, 5))


    # Fill descriptions from descriptions sheet.
    # Column 1. Name or id as convenient
    # Column 2. Group name (full path to group).
    def groups_filler(self, filename : str = "new_groups.xlsx"):
        groups_dict = {}

        for row in range(1, self.groups_sheet.max_row + 1):
            id_name = self.groups_sheet.cell(row, 1).value
            group_name = self.groups_sheet.cell(row, 2).value
            groups_dict.update([(id_name, group_name)])

        for row in range(2, self.export_sheet.max_row + 1):
            id_name = self.export_sheet.cell(row, 3).value
            if id_name in groups_dict.keys():
                group_name = groups_dict[id_name]
                self.export_sheet.cell(row, 3).value = group_name
                print(self.GREEN(f"{row}. changed"))
            else:
                print(self.YELLOW(f"{row}. skipped"))

        self.export_file.save(filename)
        print(self.GREEN(f"\nFile {filename} created"))