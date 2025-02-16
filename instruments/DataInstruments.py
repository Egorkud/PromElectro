import openpyxl
import os
from pdf2image import convert_from_path
from io import BytesIO
import img2pdf

from instruments.Resources import Resources


class DataInstruments(Resources):
    def __init__(self):
        super().__init__()

    def init_project(self):
        ...

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
