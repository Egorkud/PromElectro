import hashlib
import os
import time
import random
import re
from urllib.parse import urlparse
import requests
from pathlib2 import Path
from PyPDF2 import PdfReader
from io import BytesIO
from fake_useragent import UserAgent
import logging
import csv
from datetime import datetime

from instruments.Resources import Resources


class BaseParser(Resources):
    def __init__(self):
        super().__init__()

        self.headers = {
            "User-Agent": UserAgent().random,
        }
        self.counter = 43
        self.char_dict = {}
        self.images_counter = 0
        self.instructions_counter = 0
        self.log_data = []

    @staticmethod
    def save_html(req: object, new_name: str = "index.html") -> None:
        """
        Creates index.html file for
        :param req: requests object
        :param new_name: Name of html file (optional)
        """
        with open(f"{new_name}", "w", encoding="utf-8") as f:
            f.write(req.text)

    # Returns title of the pdf file
    @staticmethod
    def read_pdf(request) -> str:
        pdf_data = BytesIO(request.content)
        pdf_reader = PdfReader(pdf_data)
        metadata = pdf_reader.metadata
        title = metadata.get("/Title", "Назва не знайдена")

        return title

    @staticmethod
    def save_file_with_hash(file_path: Path, request, extension, idx = "") -> str:
        """

        :param file_path: Path to file to be saved.
        :param request: Request from website.
        :param extension: .pdf or .jpg to save file.
        :param idx: Optional for photos (if there are some images for 1 item).
        :return: Returns changed name with hash.
        """
        file_content = request.content  # Отримуємо вміст файлу
        file_stem = file_path.stem  # Початкова назва без розширення

        # Генеруємо хеш (беремо 8 символів для унікальності)
        file_hash = hashlib.sha256(file_content).hexdigest()[:18]

        # Формуємо нове ім'я файлу
        new_file_name = f"{file_stem}_{idx}{file_hash}{extension}"

        new_file_path = file_path.parent / new_file_name

        # Записуємо файл
        with open(new_file_path, "wb") as file:
            file.write(file_content)

        return new_file_name

    def save_names_data(self, filename, item_type, last_name, item_articule, series, manufacturer, row, idx):
        self.names_sheet.cell(row, 1 + idx).value = item_type
        self.names_sheet.cell(row, 3 + idx).value = last_name
        self.names_sheet.cell(row, 5).value = item_articule
        self.names_sheet.cell(row, 6).value = series
        self.names_sheet.cell(row, 7).value = manufacturer

        self.book_names_data.save(f"Names_data{filename}.xlsx")

    def download_instruction_file(self, instruction_link, row):
        output_folder = "downloaded_pdfs"
        os.makedirs(output_folder, exist_ok=True)

        time.sleep(1 + random.uniform(1, 2))
        req_pdf = requests.get(instruction_link, headers=self.headers)

        title = self.read_pdf(req_pdf)
        file_name = Path(title).stem
        # Перевірка на кирилицю
        if re.search(r'[^a-zA-Z0-9_\-]', file_name):
            file_name = "Instruction_"

        file_path_no_hash = Path(output_folder) / file_name
        file_name_with_hash = self.save_file_with_hash(file_path_no_hash, req_pdf, ".pdf")

        server_file_path = f"/content/instructions/{file_name_with_hash}"
        self.instructions_counter += 1
        self.blank_sheet.cell(row, 7).value = server_file_path

    def check_key(self, key):
        if key not in self.char_dict.keys():
            self.char_dict.update([(key, self.counter)])
            self.blank_sheet.cell(1, self.counter).value = key
            self.counter += 1

    def download_photos(self, photo_links, row, folder_name):
        output_folder = "downloaded_photos"
        os.makedirs(output_folder, exist_ok=True)

        for idx, link in enumerate(photo_links):
            try:
                file_name = os.path.basename(urlparse(link).path)

                if re.search(r'[^a-zA-Z0-9_\-]', file_name):
                    file_name = f"Photo_"

                file_path_no_hash = Path(output_folder) / file_name


                photo_path_name = f"/content/images/ctproduct_image/{folder_name}"

                time.sleep(0.5 + random.uniform(1, 2))
                req = requests.get(link, headers=self.headers)

                # Використання save_file_with_hash для збереження
                file_name_with_hash = self.save_file_with_hash(file_path_no_hash, req,
                                                               ".jpg")

                self.images_counter += 1
                self.blank_sheet.cell(row, 16 + idx).value = f"{photo_path_name}/{file_name_with_hash}"
            except:
                pass


class ParserLogger:
    def __init__(self, log_name="parser"):
        """Ініціалізація логера з можливістю передавати ім'я лог-файлу"""
        self.date_str = datetime.now().strftime("%Y-%m-%d")  # Поточна дата
        self.log_name = log_name

        # Формуємо шляхи для лог-файлів
        self.log_file = f"logs/{self.log_name}_{self.date_str}.log"
        self.csv_file = f"logs/{self.log_name}_{self.date_str}.csv"

        # Переконуємося, що папка logs існує
        os.makedirs("logs", exist_ok=True)

        # Налаштовуємо логер
        self.logger = self.setup_logger()

    def setup_logger(self):
        """Налаштовує логер для запису у файл"""
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[
                logging.FileHandler(self.log_file, mode="a", encoding="utf-8"),
                logging.StreamHandler()  # Додатково виводить логи у консоль
            ]
        )
        return logging.getLogger("ParserLogger")

    def log_to_csv(self, data, header=None):
        """Записує дані у CSV-файл із роздільником ';' і додає дату та час у першу колонку"""
        file_exists = os.path.isfile(self.csv_file)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Поточний час у правильному форматі

        with open(self.csv_file, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file, delimiter=";", quotechar='"', quoting=csv.QUOTE_MINIMAL)  # ';' як роздільник

            # Якщо файл новий, додаємо заголовки
            if header and not file_exists:
                writer.writerow(["Дата та час"] + list(header))  # Додаємо заголовок для часу

            writer.writerow([timestamp] + list(data))  # Додаємо час як ОДИН елемент списку

    def log_parsing_result(self, data, header=None):
        """Логує дані у TXT і записує у CSV"""
        self.logger.info(f"Parsing: {data}")

        # Запис у CSV (дані розподіляються по колонках)
        self.log_to_csv(data, header)