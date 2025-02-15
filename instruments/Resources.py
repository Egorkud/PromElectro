from io import BytesIO

import openpyxl
import hashlib
from pathlib2 import Path
from PyPDF2 import PdfReader
from colorama import Fore, Back, Style, init


class Resources:
    def __init__(self):
        # Common data (usually does not need changes)
        try:
            self.blank_file = openpyxl.open("data/sample.xlsx")
            self.blank_sheet = self.blank_file.active
            self.book_empty = openpyxl.Workbook()  # Empty table
            self.empty_sheet = self.book_empty.active
        except FileNotFoundError as ex:
            print(ex)
            print("Problems with common data files load. Use init_project()\n")
            raise SystemExit



        # Adding colours for cosy prints
        init(autoreset=True)
        self.GREEN = lambda text: f"{Fore.GREEN}{text}{Style.RESET_ALL}"
        self.RED = lambda text: f"{Fore.RED}{text}{Style.RESET_ALL}"
        self.YELLOW = lambda text: f"{Fore.YELLOW}{text}{Style.RESET_ALL}"
        self.BLUE = lambda text: f"{Fore.BLUE}{text}{Style.RESET_ALL}"

    def close(self):
        self.blank_file.close()

    # Returns title of the pdf file
    @staticmethod
    def read_pdf(request) -> str:
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