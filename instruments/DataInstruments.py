import openpyxl
from openpyxl.workbook import Workbook
from pathlib2 import Path
import pandas as pd

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
                        print(self.GREEN("Directory '{i}' created"))

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


    # Simply gets data from one file and write to empty sheet excel
    def get_coloured_cells(self):
        counter = 1
        for row in range(1, self.work_sheet.max_row + 1):

            name = self.work_sheet.cell(row, 1).value
            articule = self.work_sheet.cell(row, 3).value
            cell_fill = self.work_sheet.cell(row, 11).fill

            if cell_fill.bgColor.rgb != "00000000":
                print(row)
                self.empty_sheet.cell(counter, 1).value = name
                self.empty_sheet.cell(counter, 2).value = articule

                counter += 1

        self.book_empty.save("new_filtered_data.xlsx")

    # Fill descriptions from descriptions sheet.
    # Column 1. Name or id as convenient
    # Column 2. Group name (full path to group).
    def groups_filler(self, filename : str = "new_groups.xlsx"):
        groups_dict = {}

        for row in range(1, self.groups_sheet.max_row + 1):
            id_name = self.groups_sheet.cell(row, 1).value
            group_name = self.groups_sheet.cell(row, 2).value
            groups_dict.update([(id_name, group_name)])

        for row in range(2, self.data_sheet.max_row + 1):
            id_name = self.data_sheet.cell(row, 3).value
            if id_name in groups_dict.keys():
                group_name = groups_dict[id_name]
                self.data_sheet.cell(row, 3).value = group_name
                print(self.GREEN(f"{row}. changed"))
            else:
                print(self.YELLOW(f"{row}. skipped"))

        self.groups_file.save(filename)
        print(self.GREEN(f"\nFile {filename} created"))

    @staticmethod
    def check_duplicates_articule(export_file : str = "name.xlsx", work_file : str = "name.xlsx"):
        # Завантажуємо дані з другого стовпця (артикули)
        export_df = pd.read_excel(export_file, usecols=[1])  # 0-based index → 2-й стовпець = index 1
        work_df = pd.read_excel(work_file, usecols=[1])

        # Конвертуємо артикули в множину для швидкого пошуку
        export_articles = set(export_df.iloc[:, 0].dropna())

        # Перевіряємо наявність у множині
        duplicates = work_df.iloc[:, 0].dropna().isin(export_articles)

        # Виводимо рядки з дублями
        for idx, is_duplicate in enumerate(duplicates, start=2):
            if is_duplicate:
                print(f"{work_df.iloc[idx-2, 0]}: {idx}")

        print("✅ check_duplicates_articule Done!")