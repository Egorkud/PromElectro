import openpyxl
import os
import pandas as pd
from openpyxl.pivot.fields import Boolean
from openpyxl.workbook import Workbook
from pathlib2 import Path
from pdf2image import convert_from_path
from io import BytesIO
import img2pdf

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
            sheet.title = "Sheet"

            for id, name in config.PRODUCT_COLUMNS.items():
                sheet.cell(1, id).value = name

            wb.save(sample_file)
            print(self.GREEN(f"File {sample_file} was filled"))

        print(self.BLUE("\nProject initialisation finished\n"))

    # Fill descriptions from descriptions sheet.
    # Column 1. Name or id as convenient
    # Column 2. Group name (full path to group).
    def groups_filler(self,
                      filename:str = "new_groups.xlsx",
                      export_file:str = "export.xlsx"):
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


    # Compress all the files by screenshotting pages
    @staticmethod
    def compress_pdf_folder(input_folder:str = "downloaded_pdfs",
                            output_folder:str = "compressed_pdfs",
                            dpi:int = 200):
        # Потрібно завантажити цей інструмент та можна просто додати в
        # директорію проєкту та далі вказати тут шлях до нього
        poppler_path = r"../data/poppler-24.08.0/Library/bin"

        # Створюємо вихідну папку, якщо її немає
        os.makedirs(output_folder, exist_ok=True)

        # Перебираємо всі файли в папці
        for file_name in os.listdir(input_folder):
            if file_name.lower().endswith(".pdf"):
                input_pdf = os.path.join(input_folder, file_name)
                output_pdf = os.path.join(output_folder, file_name)

                images = convert_from_path(input_pdf, dpi=dpi, poppler_path=poppler_path)

                img_bytes = []
                for img in images:
                    img_buffer = BytesIO()
                    img.save(img_buffer, format="JPEG", quality=50)  # Стиснення JPEG
                    img_bytes.append(img_buffer.getvalue())

                with open(output_pdf, "wb") as f:
                    f.write(img2pdf.convert(img_bytes))

                print(f"Стиснуто: {file_name}")


    # Можна поміняти файли місцями, щоб перевірити роботу про всяк випадок
    # Для цього додано параметр reverse_check
    @staticmethod
    def check_duplicates_articule(export_file:str = "name.xlsx",
                                  work_file:str = "name.xlsx",
                                  reverse_check:bool = False):

        # Завантажуємо дані з обраного стовпця (артикули)
        export_df = pd.read_excel(export_file, usecols=[1])  # 0-based index → 2-й стовпець = index 1
        work_df = pd.read_excel(work_file, usecols=[1])

        # Розвертаємо перевірку файлів у іншу сторону
        if reverse_check:
            export_df, work_df = work_df, export_df

        # Конвертуємо артикули в множину для швидкого пошуку
        export_articles = set(export_df.iloc[:, 0].dropna())

        # Перевіряємо наявність у множині
        duplicates = work_df.iloc[:, 0].dropna().isin(export_articles)

        # Виводимо рядки з дублями
        for idx, is_duplicate in enumerate(duplicates, start=2):
            if is_duplicate:
                print(f"{work_df.iloc[idx-2, 0]}: {idx}")

        print("Check for duplicates done!")


    # Simply gets data from one file and write to empty sheet excel
    def get_coloured_cells(self, file_name:str = "name.xlsx"):
        counter = 1
        work_sheet = openpyxl.open(file_name).active
        for row in range(1, work_sheet.max_row + 1):

            name = work_sheet.cell(row, 1).value
            articule = work_sheet.cell(row, 3).value
            cell_fill = work_sheet.cell(row, 11).fill

            # Checks whether cell is coloured
            if cell_fill.bgColor.rgb != "00000000":
                print(row)
                self.empty_sheet.cell(counter, 1).value = name
                self.empty_sheet.cell(counter, 2).value = articule

                counter += 1

        self.book_empty.save("new_filtered_data.xlsx")