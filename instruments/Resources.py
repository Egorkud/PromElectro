from io import BytesIO

import openpyxl
import hashlib
from pathlib2 import Path
from PyPDF2 import PdfReader
from colorama import Fore, Back, Style, init
from instruments import config


class Resources:
    def __init__(self):

        # work file load
        try:
            self.work_file = openpyxl.load_workbook("new_filtered_data.xlsx")
            self.work_sheet = self.work_file.active
        except Exception as ex:
            print(ex)
            print("Problems with work_file load\n")

        # export file load
        try:
            self.groups_file = openpyxl.load_workbook("add_groups.xlsx")
            self.data_sheet = self.groups_file["Data"]
            self.groups_sheet = self.groups_file["Groups"]
        except Exception as ex:
            print(ex)
            print("Problems with add_groups.xlsx load\n")


        # Common data (usually does not need changes)
        try:
            self.blank_file = openpyxl.open("data/sample.xlsx")
            self.blank_sheet = self.blank_file.active
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
            self.groups_file.close()
            self.blank_file.close()
            self.work_file.close()
        except Exception as ex:
            print(ex)
            print("Cannot close files, chech all the excel files or Resources.py\n")

    # Returns title of the pdf file
    @staticmethod
    def read_pdf(request):
        pdf_data = BytesIO(request.content)
        pdf_reader = PdfReader(pdf_data)
        metadata = pdf_reader.metadata
        title = metadata.get("/Title", "Назва не знайдена")

        return title

    @staticmethod
    def save_pdf(file_path: Path, request) -> str:
        file_content = request.content  # Отримуємо вміст файлу
        file_stem = file_path.stem  # Початкова назва без розширення
        file_ext = file_path.suffix  # Розширення файлу (.pdf)

        # Генеруємо хеш (беремо 8 символів для унікальності)
        file_hash = hashlib.sha256(file_content).hexdigest()[:8]

        # Формуємо нове ім'я файлу
        new_file_name = f"{file_stem}_{file_hash}{file_ext}"
        new_file_path = file_path.parent / new_file_name

        # Записуємо файл
        with open(new_file_path, "wb") as file:
            file.write(file_content)

        return new_file_name


    @staticmethod
    def clean_pdf_name(title : str):
        clean_extentions = config.clean_extentions

        try:
            title = title.split(".")[0].strip()
            if title.split(".")[1] in clean_extentions:
                return f"{title}.pdf"
            return f"{title}.pdf"
        except:
            return f"{title}.pdf"