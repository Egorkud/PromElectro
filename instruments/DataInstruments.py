import openpyxl
from openpyxl import load_workbook
import os
import pandas as pd
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

        def create_excel_file(file_path, columns):
            """Створює Excel-файл з вказаними заголовками, якщо він не існує."""
            if file_path.exists():
                return

            create_path(file_path.parent, file_path)

            wb = Workbook()
            sheet = wb.active
            sheet.title = "Sheet"

            for col_id, col_name in columns.items():
                sheet.cell(1, col_id).value = col_name

            wb.save(file_path)
            print(self.GREEN(f"File {file_path} was filled"))

        print(self.BLUE("\nProject initialisation started\n"))
        # Crete folders (convenience purpose)
        folders = ("import_done", "import_queue", "temp_old")
        create_path(*(Path(i) for i in folders))

        # Create data directory and files inside
        data_dir = Path("data")
        sample_file = data_dir / "sample.xlsx"
        names_data_file = data_dir / "names_data.xlsx"

        create_excel_file(sample_file, config.SAMPLE_PRODUCT_COLUMNS)
        create_excel_file(names_data_file, config.NAMES_DATA_COLUMNS)

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
    def check_duplicates_articule(export_file: str = "name.xlsx",
                                  work_file: str = "name.xlsx",
                                  duplicates_file: str = "duplicates.xlsx",
                                  unique_file: str = "unique.xlsx",
                                  export_cols: tuple = (1, 2),  # (артикул, назва товару)
                                  work_cols: tuple = (1, 2)):

        # Завантажуємо файли через pandas
        export_df = pd.read_excel(export_file, usecols=list(export_cols))
        work_df = pd.read_excel(work_file, usecols=list(work_cols))

        # Завантажуємо оригінальний файл через openpyxl (для отримання гіперпосилань)
        wb = load_workbook(work_file, data_only=True)
        ws = wb.active

        # Конвертуємо артикули в множину для швидкого пошуку (з .strip())
        export_articles = set(export_df.iloc[:, 0].dropna().astype(str).str.strip())

        # Фільтруємо дублікати та унікальні записи
        work_df = work_df.dropna(subset=[work_df.columns[0]])
        work_df["Артикул"] = work_df.iloc[:, 0].astype(str).str.strip()
        work_df["Назва"] = work_df.iloc[:, 1]

        # Отримуємо URL із гіпертексту у стовпці "Назва товару"
        work_df["Посилання"] = None
        for row_idx in range(2, ws.max_row + 1):  # Починаємо з 2-го рядка (1-й = заголовки)
            cell = ws.cell(row=row_idx, column=work_cols[1] + 1)  # Колонка з назвами товарів
            if cell.hyperlink:
                work_df.at[row_idx - 2, "Посилання"] = cell.hyperlink.target  # Зберігаємо URL

        # Визначаємо дублікати
        work_df["Дублікат"] = work_df["Артикул"].isin(export_articles)
        duplicates = work_df[work_df["Дублікат"]][["Артикул", "Назва", "Посилання"]]
        unique = work_df[~work_df["Дублікат"]][["Артикул", "Назва", "Посилання"]]

        # Функція для збереження у файл
        def save_to_excel(df, filename, sheet_name):
            wb_out = Workbook()
            ws_out = wb_out.active
            ws_out.title = sheet_name

            # Записуємо заголовки
            headers = ["Артикул", "Назва товару", "Посилання"]
            ws_out.append(headers)

            # Записуємо дані
            for row in df.itertuples(index=False):
                ws_out.append(list(row))

            # Зберігаємо файл
            wb_out.save(filename)

        # Збереження дублікатів
        save_to_excel(duplicates, duplicates_file, "Дублікати")

        # Збереження унікальних товарів
        save_to_excel(unique, unique_file, "Унікальні")

        print(f"✅ Дублікати збережено у '{duplicates_file}'")
        print(f"✅ Унікальні товари збережено у '{unique_file}'")


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

    @staticmethod
    def name_changer(input_file:str = "import_queue/Освещение=Светильники.xlsm",
                     output_file:str = "import_queue/New_names.xlsx",
                     col_name1:int = 4,
                     col_name2:int = 5,
                     col_article:int = 2,
                     col_group:int = 3,
                     selected_group:str = "Освещение=>Светильники=>Бытовые светильники=>Точечные светильники",
                     attribute_cols:set = (109, 100, 77)):
        """
           Оновлює дві колонки назв, додаючи перед артикулом дані з вказаних колонок,
           але тільки для товарів, що належать до заданої групи.
           Рядки, які не підходять, пропускаються.

           :param input_file: Назва вхідного файлу Excel
           :param output_file: Назва вихідного файлу Excel
           :param col_name1: Номер першої колонки з назвами (1-індекс)
           :param col_name2: Номер другої колонки з назвами (1-індекс)
           :param col_article: Номер колонки з артикулами (1-індекс)
           :param col_group: Номер колонки з групою товарів (1-індекс)
           :param selected_group: Назва групи, за якою фільтруємо
           :param attribute_cols: Номери колонок (1-індекс), дані яких вставлятимуться в назву
        """

        # Завантажуємо дані
        df = pd.read_excel(input_file, dtype=str)

        # Конвертуємо індексацію (Excel → Python, тобто віднімаємо 1)
        col_name1 -= 1
        col_name2 -= 1
        col_article -= 1
        col_group -= 1
        attribute_cols = {col - 1 for col in attribute_cols}

        # Отримуємо назви колонок
        name_col1 = df.columns[col_name1]
        name_col2 = df.columns[col_name2]
        article_col = df.columns[col_article]
        group_col = df.columns[col_group]
        attribute_columns = [df.columns[i] for i in attribute_cols]

        def modify_name(row, name_col):
            name = row[name_col]
            article = str(row[article_col])

            # Збираємо атрибути
            attributes = " ".join(str(row[col]) for col in attribute_columns if pd.notna(row[col]))

            # Якщо артикул є у назві — вставляємо перед ним атрибути
            if article in name:
                return name.replace(article, attributes + " " + article)

            return name  # Якщо артикул не знайдено, залишаємо без змін

        # Фільтруємо рядки тільки для обраної групи
        df_filtered = df[df[group_col] == selected_group].copy()

        # Оновлюємо обидві колонки назв
        df_filtered["Оновлена Назва (RU)"] = df_filtered.apply(lambda row: modify_name(row, name_col1), axis=1)
        df_filtered["Оновлена Назва (UA)"] = df_filtered.apply(lambda row: modify_name(row, name_col2), axis=1)

        # Зберігаємо тільки артикул + оновлені назви
        df_result = df_filtered[[df.columns[col_article], "Оновлена Назва (RU)", "Оновлена Назва (UA)"]]

        # Зберігаємо результат без пустих рядків
        df_result.to_excel(output_file, index=False, engine="openpyxl")

        print(f"✅ Файл збережено: {output_file}, збережено {len(df_result)} рядків")

    @staticmethod
    def get_code_row(input_file: str = "website_positions.xlsx",
                     output_file: str = "output_codes.txt"):

        # Зчитуємо тільки 1-й стовпець (index 0) без заголовків
        df = pd.read_excel(input_file, usecols=[0], dtype=str, engine='openpyxl')

        # Видаляємо порожні значення та дублікати
        unique_codes = df.iloc[:, 0].dropna().unique()

        # Швидкий запис у файл через кому
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(",".join(unique_codes))

        print(f"Збережено {len(unique_codes)} кодів у {output_file}")
