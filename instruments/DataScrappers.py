from bs4 import BeautifulSoup
import os
import requests
import time
import random
from pathlib2 import Path
from urllib.parse import urlparse
from fake_useragent import UserAgent


from instruments.Resources import Resources



class DataScrappers(Resources):
    def __init__(self):
        super().__init__()

        self.headers = {
            "User-Agent": UserAgent().random,
        }


    def instructions_from_links(self, filename : str = "scrapped_data (instructions).xlsx"):
        for row in range(2, self.export_sheet.max_row + 1):
            print(row)
            url = self.export_sheet.cell(row, 1).value

            output_folder = "downloaded_pdfs"
            os.makedirs(output_folder, exist_ok=True)

            full_url = url.split(".pdf")
            link_name, page_num = full_url[0], full_url[1]

            clean_pdf_url = f"{link_name}.pdf"

            file_name = os.path.basename(urlparse(clean_pdf_url).path)
            file_path = Path(output_folder) / file_name

            print(file_name)

            instruction_name = f"/content/instructions/{file_name}{link_name[1]}"
            self.export_sheet.cell(row, 5).value = instruction_name

            try:
                req = requests.get(url, headers=self.headers)

                # Implement data scrapping here


                # Save data to file
                with open(file_path, "wb") as pdf_file:
                    for chunk in req.iter_content(chunk_size=1024):
                        pdf_file.write(chunk)

            except Exception as ex:
                print(ex)

            self.export_file.save(filename)
            time.sleep(3 + random.uniform(1, 3))
        print(self.GREEN(f"\nFile {filename} created"))

    def photo_from_urls(self, filename : str = "scrapped_data (photos).xlsx"):
        for row in range(2, self.export_sheet.max_row + 1):
            print(self.GREEN(f"{row}. Started"))
            url = self.export_sheet.cell(row, 16).value

            output_folder = "downloaded_photos"
            os.makedirs(output_folder, exist_ok=True)

            file_path_name = os.path.basename(urlparse(url).path)
            file_path = os.path.join(output_folder, file_path_name)
            photo_path_name = f"/content/images/ctproduct_image/INPUT_MANUFACTURER_NAME"    # Needs INTPUT

            try:
                req = requests.get(url, headers=self.headers)
                src = req.text
                soup = BeautifulSoup(src, "lxml")

                # Get all the necessary links from website into list
                # If there are more than 1 photo, make 1 list with all photos
                main_photo = [...]
                all_photos = [...]


                # INPUT list of all links or only 1 main image link
                for id, link in enumerate(...):
                    req = requests.get(link, headers=self.headers)
                    file_name = f"{file_path_name}_{id + 1}.jpg"

                    with open(f"{file_path}_{id + 1}.jpg", "wb") as file:
                        file.write(req.content)

                    print(f"{id + 1}. Downloaded {link}")
                    time.sleep(3 + random.uniform(1, 3))

                    self.export_sheet.cell(row, 16 + id).value = f"{photo_path_name}/{file_name}"

            except Exception as ex:
                print(ex)

            self.export_file.save(filename)
            time.sleep(3 + random.uniform(1, 3))
        print(self.GREEN(f"\nFile {filename} created"))

    def characteristics_from_articules(self, filename : str = "scrapped_data (characteristics).xlsx"):
        char_dict = {}
        counter = 43

        for row in range(2, self.export_sheet.max_row + 1):
            print(self.GREEN(f"{row}. Started"))

            link = ""
            url_search = ""

            item_articule = self.export_sheet.cell(row, 2).value
            item_articule.replace(" ", "+")
            url = f"{url_search}{item_articule}"

            # 1 Layer
            try:
                req = requests.get(url, headers=self.headers)
                src = req.text
                soup = BeautifulSoup(src, "lxml")

                # Implement soup find searched item
                searched_item = ...

                link = f"{link}{searched_item}"
                print(link)

                time.sleep(2 + random.uniform(1, 3))
                # 2 Layer
                try:
                    req = requests.get(link, headers=self.headers)
                    src = req.text
                    soup = BeautifulSoup(src, "lxml")

                    # Continue Implementation


                except Exception as ex:
                    print(self.RED(ex))
            except Exception as ex:
                print(self.RED(ex))

            self.export_file.save(filename)
            time.sleep(3 + random.uniform(1, 3))
        print(self.GREEN(f"\nFile {filename} created"))