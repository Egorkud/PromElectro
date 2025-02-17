import hashlib
import os
import time
import random
from urllib.parse import urlparse
import requests
from pathlib2 import Path
from PyPDF2 import PdfReader
from io import BytesIO
from fake_useragent import UserAgent

from instruments.Resources import Resources


class BaseScrapper(Resources):
    def __init__(self):
        super().__init__()

        self.headers = {
            "User-Agent": UserAgent().random,
        }


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

        # Генеруємо хеш (беремо 8 символів для унікальності)
        file_hash = hashlib.sha256(file_content).hexdigest()[:8]

        # Формуємо нове ім'я файлу
        new_file_name = f"{file_stem}_{file_hash}.pdf"
        new_file_path = file_path.parent / new_file_name

        # Записуємо файл
        with open(new_file_path, "wb") as file:
            file.write(file_content)

        return new_file_name

    def save_names_data(self, item_type, last_name, item_articule, series, manufacturer, row, idx):
        self.names_sheet.cell(row, 1 + idx).value = item_type
        self.names_sheet.cell(row, 3 + idx).value = last_name
        self.names_sheet.cell(row, 5).value = item_articule
        self.names_sheet.cell(row, 6).value = series
        self.names_sheet.cell(row, 7).value = manufacturer

        self.book_names_data.save("Names_data.xlsx")

    def download_instruction_file(self, instruction_link, row):
        output_folder = "downloaded_pdfs"
        os.makedirs(output_folder, exist_ok=True)

        time.sleep(1 + random.uniform(1, 2))
        req_pdf = requests.get(instruction_link, headers=self.headers)
        title = self.read_pdf(req_pdf)
        file_name = Path(title).stem
        file_path_no_hash = Path(output_folder) / file_name

        file_name_with_hash = self.save_pdf(file_path_no_hash, req_pdf)

        server_file_path = f"/content/instructions/{file_name_with_hash}"
        self.blank_sheet.cell(row, 7).value = server_file_path

    def check_key(self, key, char_dict, counter):
        if key not in char_dict.keys():
            char_dict.update([(key, counter)])
            self.blank_sheet.cell(1, counter).value = key
            counter += 1

        return counter, char_dict

    def download_photos(self, photo_links, row):
        for id, link in enumerate(photo_links):
            try:
                output_folder = "downloaded_photos"
                os.makedirs(output_folder, exist_ok=True)

                file_path_name = os.path.basename(urlparse(link).path)
                file_path_no_hash = os.path.join(output_folder, file_path_name)
                photo_path_name = f"/content/images/ctproduct_image/FOLDER_NAME"  # Needs INTPUT
                file_name = f"{file_path_name}_{id + 1}.jpg"

                time.sleep(0.5 + random.uniform(1, 2))
                req = requests.get(link, headers=self.headers)
                with open(f"{file_path_no_hash}_{id + 1}.jpg", "wb") as file:
                    file.write(req.content)

                self.blank_sheet.cell(row, 16 + id).value = f"{photo_path_name}/{file_name}"
            except:
                pass