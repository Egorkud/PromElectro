from openpyxl.workbook import Workbook
from pathlib2 import Path
import win32com.client

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

        pult_file = data_dir / "Пульт.xlsm"

        if not pult_file.exists() or True:
            # 1. Create Excel
            absolut = Path(pult_file).resolve()
            print(absolut)
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False

            # 2 Create .xlsm file
            wb = excel.Workbooks.Add()
            wb.SaveAs(r"D:\Programming\PythonProjects\PromElectro\data\Пульт.xlsm", FileFormat=52)  # 52 = .xlsm
            wb.Close()
            excel.Quit()

            # 3. Add VBA macro
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Робимо Excel невидимим
            wb_macro = excel.Workbooks.Open(absolut)

            # Додаємо модуль VBA
            vb_component = wb_macro.VBProject.VBComponents.Add(1)  # 1 - стандартний модуль
            vb_component.Name = "Module1"
            vb_component.CodeModule.AddFromString('''
            Sub HelloWorld()
                MsgBox "Hello from VBA!", vbInformation, "Python VBA Macro"
            End Sub
            ''')

            # 5️⃣ ВАЖЛИВО! Спочатку зберігаємо макрос, і лише потім додаємо кнопку
            wb_macro.Save()

            # 6️⃣ Додаємо кнопку
            sheet = wb_macro.Sheets("Sheet1")

            excel_button = sheet.OLEObjects().Add(
                ClassType="Forms.CommandButton.1",
                Left=100, Top=100, Width=100, Height=30
            )
            excel_button.Object.Caption = "Run Macro"

            # 7️⃣ Правильний формат виклику макросу!
            excel_button.Object.OnAction = "Module1.HelloWorld"

            # 8️⃣ Фінальне збереження
            wb_macro.Save()
            wb_macro.Close()
            excel.Quit()


    # Fill descriptions from descriptions sheet.
    # Column 1. Name or id as convenient
    # Column 2. Group name (full path to group).
    def groups_filler(self, filename : str = "new_groups.xlsx"):
        groups_dict = {}

        for row in range(1, self.groups_sheet.max_row + 1):
            id_name = self.groups_sheet.cell(row, 1).value
            group_name = self.groups_sheet.cell(row, 2).value
            groups_dict.update([(id_name, group_name)])

        for row in range(2, self.export_sheet.max_row + 1):
            id_name = self.export_sheet.cell(row, 3).value
            if id_name in groups_dict.keys():
                group_name = groups_dict[id_name]
                self.export_sheet.cell(row, 3).value = group_name
                print(self.GREEN(f"{row}. changed"))
            else:
                print(self.YELLOW(f"{row}. skipped"))

        self.export_file.save(filename)
        print(self.GREEN(f"\nFile {filename} created"))