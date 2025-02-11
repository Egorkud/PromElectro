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

    def big_scrap_from_articules(self, filename : str = "scrapped_data (instructions+photos).xlsx"):
        char_dict = {}
        counter = 43

        for row in range(89, self.work_sheet.max_row + 1):
            print(self.GREEN(f"{row}. Started"))

            url_search_ru = "https://feron.ua/search/?search="

            item_articule = self.work_sheet.cell(row, 2).value
            self.blank_sheet.cell(row, 2).value = item_articule

            url = f"{url_search_ru}{item_articule}"

            # 1 Layer - Search
            try:
                req = requests.get(url, headers=self.headers)
                src = req.text
                soup = BeautifulSoup(src, "lxml")

                # Implement soup find searched item
                searched_item = (soup.find("div", class_="image")
                                 .find("a")
                                 .get("href"))
                print(searched_item)
                time.sleep(2 + random.uniform(1, 3))

                # 2 Layer - Item
                try:
                    # Continue Implementation
                    req = requests.get(searched_item, headers=self.headers)
                    src = req.text
                    soup = BeautifulSoup(src, "lxml")

                    # # Get description RU
                    # try:
                    #     description_lines_ru = (soup.find("div", class_="product_tab_content")
                    #                    .find("div", id="tab-description").contents)
                    #     clean_description_ru = [str(i) for i in description_lines_ru if i != "\n"]
                    #     self.blank_sheet.cell(row, 11).value = "\n".join(clean_description_ru)
                    #
                    #     time.sleep(2 + random.uniform(1, 3))
                    # except Exception as ex:
                    #     print(ex)
                    #     print("No description RU")
                    #
                    # # Get description UKR
                    # try:
                    #     ukr_link = searched_item.replace("feron.ua/", "feron.ua/ua/")
                    #     req = requests.get(ukr_link, headers=self.headers)
                    #     src = req.text
                    #     soup_ukr = BeautifulSoup(src, "lxml")
                    #
                    #     description_lines_ukr = (soup_ukr.find("div", class_="product_tab_content")
                    #                    .find("div", id="tab-description").contents)
                    #     clean_description_ukr = [str(i) for i in description_lines_ukr if i != "\n"]
                    #     self.blank_sheet.cell(row, 12).value = "\n".join(clean_description_ukr)
                    #
                    #     time.sleep(2 + random.uniform(1, 3))
                    # except Exception as ex:
                    #     print(ex)
                    #     print("No description ukr")
                    #
                    #
                    # # Get characteristics
                    # try:
                    #     characteristics = (soup.find("div", class_="product_tab_content")
                    #                        .find("table", class_="table table-bordered")
                    #                        .find_all("tbody"))
                    #
                    #     for char in characteristics:
                    #         for tr in char.find_all("tr"):
                    #             tds = tr.find_all("td")
                    #             key, value = tds[0].text, tds[1].text
                    #
                    #             if key not in char_dict.keys():
                    #                 char_dict.update([(key, counter)])
                    #                 self.blank_sheet.cell(1, counter).value = key
                    #                 counter += 1
                    #
                    #             char_col = char_dict[key]
                    #             self.blank_sheet.cell(row, char_col).value = value
                    #
                    # except Exception as ex:
                    #     print(ex)
                    #     print("No characteristics")
                    #
                    # Get instructions
                    try:
                        find_instr = (soup.find("ul", class_="attribute attribute--insert qq3")
                                       .find_all("li"))

                        for instruction in find_instr:
                            span = instruction.find("span")
                            if span and "Инструкция" in span.text:
                                instructioin_link = instruction.find("a").get("href")

                                output_folder = "downloaded_pdfs"
                                os.makedirs(output_folder, exist_ok=True)

                                time.sleep(2 + random.uniform(1, 3))
                                req_pdf = requests.get(instructioin_link, headers=self.headers)
                                title = self.read_pdf(req_pdf)
                                file_name = self.clean_pdf_name(title)
                                file_path = Path(output_folder) / file_name

                                new_file_name = self.save_pdf(file_path, req_pdf)

                                server_file_path = f"/content/instructions/{new_file_name}"
                                self.blank_sheet.cell(row, 7).value = server_file_path


                    except Exception as ex:
                        print(ex)
                        print("No instructions")

                    # Get photos
                    try:
                        photo_data = (soup.find("div", class_="column_left-slider sliderss")
                                      .find("div", class_="slider-nav_prod")
                                      .find_all("img"))

                        photo_links = [link.get("href") for link in photo_data]


                        for id, link in enumerate(photo_links):
                            try:
                                output_folder = "downloaded_photos"
                                os.makedirs(output_folder, exist_ok=True)

                                file_path_name = os.path.basename(urlparse(link).path)
                                file_path = os.path.join(output_folder, file_path_name)
                                photo_path_name = f"/content/images/ctproduct_image/feron"  # Needs INTPUT
                                file_name = f"{file_path_name}_{id + 1}.jpg"

                                time.sleep(0.5 + random.uniform(1, 2))
                                req = requests.get(link, headers=self.headers)
                                with open(f"{file_path}_{id + 1}.jpg", "wb") as file:
                                    file.write(req.content)

                                self.blank_sheet.cell(row, 16 + id).value = f"{photo_path_name}/{file_name}"
                            except:
                                pass

                    except Exception as ex:
                        print(ex)
                        print("No photos")

                    # # Get names ru ukr
                    # try:
                    #     manufacturers = ("Ardero ", "Feron ", "Ledcoin ")
                    #     series = self.work_sheet.cell(row, 3).value
                    #
                    #     ukr_link = searched_item.replace("feron.ua/", "feron.ua/ua/")
                    #     req_ukr = requests.get(ukr_link, headers=self.headers)
                    #     src_ukr = req_ukr.text
                    #     soup_ukr = BeautifulSoup(src_ukr, "lxml")
                    #
                    #     if series is None:
                    #         series = ""
                    #     else:
                    #         series = f"{series} "
                    #
                    #     full_name_ru = (soup.find("div", class_="container product_page")
                    #                  .find("h1", itemprop="name")
                    #                  .text.strip())
                    #     full_name_ukr = (soup_ukr.find("div", class_="container product_page")
                    #                  .find("h1", itemprop="name")
                    #                  .text.strip())
                    #
                    #
                    #     manufacturer = next((m for m in manufacturers if m in full_name_ru), None)
                    #     item_type_ru, last_name_ru = full_name_ru.split(manufacturer, 1)
                    #     item_type_ukr, last_name_ukr = full_name_ukr.split(manufacturer, 1)
                    #
                    #
                    #     try:
                    #         last_name_ru = last_name_ru.replace(f"{series}", "")
                    #         last_name_ukr = last_name_ukr.replace(f"{series}", "")
                    #     except Exception as ex:
                    #         print(ex)
                    #
                    #
                    #     new_name_ru = f"{item_type_ru.strip()} {last_name_ru.strip()} {item_articule} {series}{manufacturer}"
                    #     new_name_ukr = f"{item_type_ukr.strip()} {last_name_ukr.strip()} {item_articule} {series}{manufacturer}"
                    #
                    #     print(new_name_ru)
                    #     print(new_name_ukr)
                    #
                    #     self.blank_sheet.cell(row, 4).value = new_name_ru
                    #     self.blank_sheet.cell(row, 5).value = new_name_ukr
                    #
                    # except Exception as ex:
                    #     print(ex)
                    #     print("No name or manufacturer name")


                    time.sleep(2 + random.uniform(1, 3))
                except Exception as ex:
                    print(self.RED(ex))
            except Exception as ex:
                print(self.RED(ex))

            self.blank_file.save(filename)
            time.sleep(3 + random.uniform(1, 3))
        print(self.GREEN(f"\nFile {filename} created"))