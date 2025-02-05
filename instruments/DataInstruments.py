from openpyxl.workbook import Workbook
from pathlib2 import Path
import win32com.client

from instruments import config
from instruments.Resources import Resources


class DataInstruments(Resources):
    def __init__(self):
        super().__init__()

    def init_project(self):
        def create_path(*files):
            for i in files:
                if not i.exists():
                    if i.suffix:
                        i.touch(exist_ok=True)
                        print(self.GREEN(f"File '{i}' created"))
                    else:
                        i.mkdir(exist_ok=True)
                        print(self.GREEN(f"Directory '{i}' created"))

        print(self.BLUE("\nProject initialisation started\n"))
        # Crete folders (convenience purpose)
        folders = ("import_done", "import_queue", "temp_old")
        create_path(*(Path(i) for i in folders))

        # Create data directory and files inside
        data_dir = Path("data")
        sample_file = data_dir / "sample.xlsx"

        if not sample_file.exists():
            create_path(data_dir, sample_file)

            wb = Workbook()
            sheet = wb.active
            sheet.title = "Sheet1"

            for id, name in config.PRODUCT_COLUMNS.items():
                sheet.cell(1, id).value = name

            wb.save(sample_file)
            print(self.GREEN(f"File {sample_file} was filled"))

        print(self.BLUE("\nProject initialisation finished\n"))

    # Fill descriptions from descriptions sheet.
    # Column 1. Name or id as convenient
    # Column 2. Group name (full path to group).
    def groups_filler(self, filename : str = "new_groups.xlsx"):
        groups_dict = {}
        export_sheet = self.work_file["export_sheet"]
        groups_sheet = self.work_file["groups_sheet"]

        # Get all possible groups
        for row in range(1, groups_sheet.max_row + 1):
            id_name = groups_sheet.cell(row, 1).value
            group_name = groups_sheet.cell(row, 2).value
            groups_dict.update([(id_name, group_name)])

        #
        for row in range(2, export_sheet.max_row + 1):
            id_name = export_sheet.cell(row, 3).value
            if id_name in groups_dict.keys():
                group_name = groups_dict[id_name]
                export_sheet.cell(row, 3).value = group_name
                print(self.GREEN(f"{row}. changed"))
            else:
                print(self.YELLOW(f"{row}. skipped"))

        self.work_file.save(filename)
        print(self.GREEN(f"\nFile {filename} created"))