import sys
import pandas as pd
import json
import os
import re
import datetime
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QProgressBar, QLineEdit, QComboBox,
    QPlainTextEdit, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QSpinBox, QMessageBox, QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, QEvent
from PyQt5.QtGui import QIcon

LOG_FILE = "rewriter_log.txt"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

CONFIG_FILE = "api_config.json"
PROMPTS_FILE = "prompts.json"
VARIABLES_FILE = "variables.json"

def save_api_key(api_key):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"api_key": api_key}, f)
    logging.info("API-ключ сохранён в %s", CONFIG_FILE)

def load_api_key():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            data = json.load(f)
            api_key = data.get("api_key", "")
            logging.info("Загружен API-ключ из %s", CONFIG_FILE)
            return api_key
    logging.warning("Файл %s не найден, возвращён пустой ключ", CONFIG_FILE)
    return ""

def clean_response(text):
    text = re.sub(r'```html|```', '', text).strip()
    text = re.sub(r'\n+', '\n', text)
    return text

class Worker(QObject):
    """
    Объект-«работник», который запускается в отдельном QThread,
    чтобы не блокировать UI. Он параллельно обрабатывает строки (ThreadPool),
    и после обработки каждой строки отправляет сигнал partialResult,
    чтобы UI мог «дописывать» Excel или обновлять прогресс.
    """
    partialResult = pyqtSignal(int, dict)  # row_index, данные
    progressSignal = pyqtSignal(int)       # % прогресса
    finishedSignal = pyqtSignal(str)       # итоговое сообщение или путь к файлу

    def __init__(self, file_path, api_key, model, prompt1, prompt2, variables_dict, parallel_count, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.api_key = api_key
        self.model = model
        self.prompt1 = prompt1
        self.prompt2 = prompt2
        self.variables_dict = variables_dict
        self.parallel_count = parallel_count

    def run(self):
        """
        Основной метод: считываем Excel, обрабатываем строки параллельно,
        после каждой строки отправляем partialResult, и дополнительно - progressSignal.
        """
        import openai

        # Читаем Excel
        try:
            df = pd.read_excel(self.file_path, header=0)
        except Exception as e:
            msg = f"Ошибка чтения Excel: {e}"
            logging.error(msg)
            self.finishedSignal.emit(msg)
            return

        if len(df) == 0:
            msg = "Файл не содержит строк."
            logging.warning(msg)
            self.finishedSignal.emit(msg)
            return

        # Находим, какие переменные реально нужны
        used_vars = []
        for var_name in self.variables_dict.keys():
            if f"{{{var_name}}}" in self.prompt1:
                used_vars.append(var_name)

        do_second = bool(self.prompt2.strip())

        # Формируем список колонок
        columns = used_vars + ["OpenAI_Response"]
        if do_second:
            columns.append("OpenAI_Response_2")

        # Создаём пустой Excel (или очищаем), чтобы файл уже существовал
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.out_path = self.file_path.replace(".xlsx", f"_partial_{self.model}_{timestamp}.xlsx")
        empty_df = pd.DataFrame(columns=columns)
        empty_df.to_excel(self.out_path, index=False)
        logging.info("Создан пустой файл для результатов: %s", self.out_path)

        total_rows = len(df)
        done_count = 0

        # worker-функция для одной строки
        def handle_one_row(row_index):
            row_data = df.iloc[row_index]

            # Собираем переменные
            var_values = {}
            for v in used_vars:
                col_index = self.variables_dict[v]
                val = row_data.iloc[col_index]
                val = "" if pd.isna(val) else str(val)
                var_values[v] = val

            # Подставляем в 1-й промпт
            p1 = self.prompt1
            for k, val in var_values.items():
                p1 = p1.replace(f"{{{k}}}", val)

            # 1-й запрос
            res1 = ""
            try:
                openai.api_key = self.api_key
                resp = openai.ChatCompletion.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": "Ты профессиональный копирайтер."},
                        {"role": "user", "content": p1}
                    ]
                )
                res1 = clean_response(resp["choices"][0]["message"]["content"])
            except Exception as e:
                res1 = f"Ошибка: {e}"
                logging.error("Ошибка в 1-м запросе, строка %d: %s", row_index, e)

            # 2-й запрос (если нужно)
            res2 = None
            if do_second:
                p2 = self.prompt2.replace("{prev_result}", res1)
                try:
                    resp2 = openai.ChatCompletion.create(
                        model=self.model,
                        messages=[
                            {"role": "system", "content": "Ты профессиональный копирайтер и переводчик."},
                            {"role": "user", "content": p2}
                        ]
                    )
                    res2 = clean_response(resp2["choices"][0]["message"]["content"])
                except Exception as e:
                    res2 = f"Ошибка: {e}"
                    logging.error("Ошибка во 2-м запросе, строка %d: %s", row_index, e)

            # Формируем словарь
            rd = {}
            for v in used_vars:
                rd[v] = var_values[v]
            rd["OpenAI_Response"] = res1
            if do_second:
                rd["OpenAI_Response_2"] = res2

            return (row_index, rd)

        # Запускаем пул потоков
        futures = []
        with ThreadPoolExecutor(max_workers=self.parallel_count) as executor:
            # Подготавливаем задания
            for i in range(total_rows):
                fut = executor.submit(handle_one_row, i)
                futures.append(fut)

            # Обрабатываем результаты по мере готовности
            for fut in as_completed(futures):
                row_index, row_data = fut.result()
                done_count += 1

                # Отправляем сигнал partialResult, чтобы UI мог «дополнить» Excel
                self.partialResult.emit(row_index, row_data)

                # Прогресс
                percent = int(done_count / total_rows * 100)
                self.progressSignal.emit(percent)

        # Когда всё, посылаем finishedSignal
        self.finishedSignal.emit(self.out_path)


class RewriterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.prompts_dict = {}
        self.variables_dict = {}
        self.last_focused_prompt = 1
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Выберите файл Excel с данными")
        layout.addWidget(self.label)
        
        self.file_button = QPushButton("Выбрать файл")
        self.file_button.clicked.connect(self.openFile)
        layout.addWidget(self.file_button)
        
        self.api_label = QLabel("Введите API-ключ OpenAI")
        layout.addWidget(self.api_label)
        
        self.api_input = QLineEdit()
        self.api_input.setEchoMode(QLineEdit.Password)
        self.api_input.setText(load_api_key())
        layout.addWidget(self.api_input)
        
        self.save_api_button = QPushButton("Сохранить ключ")
        self.save_api_button.clicked.connect(self.saveApiKey)
        layout.addWidget(self.save_api_button)
        
        self.model_label = QLabel("Выберите модель OpenAI")
        layout.addWidget(self.model_label)
        
        self.model_select = QComboBox()
        self.model_select.addItems(["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo", "o3-mini", "o1-mini"])
        layout.addWidget(self.model_select)

        # Поле для параллельности
        parallel_box = QHBoxLayout()
        self.parallel_label = QLabel("Кол-во параллельных запросов:")
        self.parallel_spin = QSpinBox()
        self.parallel_spin.setRange(1, 20)
        self.parallel_spin.setValue(1)
        parallel_box.addWidget(self.parallel_label)
        parallel_box.addWidget(self.parallel_spin)
        layout.addLayout(parallel_box)

        self.prompt_label = QLabel("Шаблон промпта (1-й запрос)")
        layout.addWidget(self.prompt_label)
        
        self.prompt_input = QPlainTextEdit()
        self.prompt_input.setPlainText("")
        layout.addWidget(self.prompt_input)
        self.prompt_input.installEventFilter(self)

        self.second_prompt_label = QLabel("Второй промпт (опционально). {prev_result} = результат первого запроса.")
        layout.addWidget(self.second_prompt_label)
        
        self.second_prompt_input = QPlainTextEdit()
        layout.addWidget(self.second_prompt_input)
        self.second_prompt_input.installEventFilter(self)

        # Блок сохранения/загрузки промпта
        save_prompt_layout = QHBoxLayout()
        self.prompt_name_label = QLabel("Имя шаблона:")
        save_prompt_layout.addWidget(self.prompt_name_label)

        self.prompt_name_input = QLineEdit()
        save_prompt_layout.addWidget(self.prompt_name_input)
        
        self.save_prompt_button = QPushButton("Сохранить шаблон")
        self.save_prompt_button.clicked.connect(self.savePrompt)
        save_prompt_layout.addWidget(self.save_prompt_button)

        layout.addLayout(save_prompt_layout)

        load_prompt_layout = QHBoxLayout()
        self.load_prompt_label = QLabel("Сохранённые шаблоны:")
        load_prompt_layout.addWidget(self.load_prompt_label)

        self.saved_prompts_combo = QComboBox()
        load_prompt_layout.addWidget(self.saved_prompts_combo)

        self.load_prompt_button = QPushButton("Загрузить")
        self.load_prompt_button.clicked.connect(self.loadSelectedPrompt)
        load_prompt_layout.addWidget(self.load_prompt_button)

        self.delete_prompt_button = QPushButton("Удалить")
        self.delete_prompt_button.clicked.connect(self.deleteSelectedPrompt)
        load_prompt_layout.addWidget(self.delete_prompt_button)

        layout.addLayout(load_prompt_layout)

        var_label = QLabel("Переменные (имя -> индекс). Нажмите + чтобы вставить:")
        layout.addWidget(var_label)

        self.variables_table = QTableWidget()
        self.variables_table.setColumnCount(3)
        self.variables_table.setHorizontalHeaderLabels(["Имя переменной", "Индекс", "Действия"])
        self.variables_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.variables_table)

        add_var_layout = QHBoxLayout()
        self.var_name_input = QLineEdit()
        self.var_name_input.setPlaceholderText("Напр. color")
        add_var_layout.addWidget(self.var_name_input)
        
        self.var_col_index = QSpinBox()
        self.var_col_index.setRange(0,999)
        add_var_layout.addWidget(self.var_col_index)

        self.add_var_button = QPushButton("Добавить переменную")
        self.add_var_button.clicked.connect(self.addVariable)
        add_var_layout.addWidget(self.add_var_button)

        layout.addLayout(add_var_layout)

        self.start_button = QPushButton("Начать обработку")
        self.start_button.clicked.connect(self.startProcessing)
        self.start_button.setEnabled(False)
        layout.addWidget(self.start_button)

        self.progress = QProgressBar()
        self.progress.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress)

        self.status_label = QLabel("Статус: ожидание файла")
        layout.addWidget(self.status_label)

        self.setLayout(layout)
        self.setWindowTitle("Обработка + построчное сохранение")
        self.resize(700,800)

        self.loadPromptsToCombo()
        self.loadVariables()

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.FocusIn:
            if obj == self.prompt_input:
                self.last_focused_prompt = 1
            elif obj == self.second_prompt_input:
                self.last_focused_prompt = 2
        return super().eventFilter(obj, ev)

    def openFile(self):
        options = QFileDialog.Options()
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel", "",
            "Excel (*.xlsx);;All Files (*)",
            options=options
        )
        if path:
            self.file_path = path
            self.label.setText(f"Файл: {path}")
            self.start_button.setEnabled(True)

    def saveApiKey(self):
        api_key = self.api_input.text().strip()
        save_api_key(api_key)
        self.status_label.setText("API-ключ сохранён.")

    def startProcessing(self):
        if not hasattr(self, "file_path"):
            self.status_label.setText("Сначала выберите файл Excel.")
            return
        api_key = self.api_input.text().strip()
        if not api_key:
            self.status_label.setText("API-ключ не указан.")
            return

        model = self.model_select.currentText()
        pr1 = self.prompt_input.toPlainText()
        pr2 = self.second_prompt_input.toPlainText()
        parallel_count = self.parallel_spin.value()

        # Создаём объект-«работник» и отдельный QThread
        self.worker_thread = QThread()
        self.worker_obj = Worker(
            file_path=self.file_path,
            api_key=api_key,
            model=model,
            prompt1=pr1,
            prompt2=pr2,
            variables_dict=self.variables_dict,
            parallel_count=parallel_count
        )
        self.worker_obj.moveToThread(self.worker_thread)

        # Подключаем сигналы
        self.worker_obj.partialResult.connect(self.handlePartialResult)
        self.worker_obj.progressSignal.connect(self.progress.setValue)
        self.worker_obj.finishedSignal.connect(self.handleFinished)

        self.worker_thread.started.connect(self.worker_obj.run)

        # Запускаем
        self.worker_thread.start()
        self.status_label.setText("Обработка запущена...")
        self.start_button.setEnabled(False)

    def handlePartialResult(self, row_index, row_dict):
        """
        Здесь мы «дополняем» Excel-файл по одной строке.
        """
        # Допишем строку в уже созданный файл
        # Файл создаётся в Worker.run() (пустой DataFrame).
        out_path = self.worker_obj.out_path  # Путь к файлу

        # 1. Считываем текущий Excel
        current_df = pd.read_excel(out_path)

        # 2. Создаём DataFrame с одной строкой
        new_df = pd.DataFrame([row_dict])

        # 3. Добавляем к current_df
        result_df = pd.concat([current_df, new_df], ignore_index=True)

        # 4. Сохраняем обратно
        result_df.to_excel(out_path, index=False)

        logging.info("Дополнена строка %d, файл: %s", row_index, out_path)

    def handleFinished(self, msg):
        self.status_label.setText(f"Готово: {msg}")
        self.start_button.setEnabled(True)
        # Останавливаем поток
        self.worker_thread.quit()
        self.worker_thread.wait()

    def loadPromptsToCombo(self):
        from pathlib import Path
        if not Path(PROMPTS_FILE).exists():
            self.prompts_dict = {}
        else:
            with open(PROMPTS_FILE,"r", encoding="utf-8") as f:
                self.prompts_dict = json.load(f)
        self.saved_prompts_combo.clear()
        for k in self.prompts_dict.keys():
            self.saved_prompts_combo.addItem(k)

    def savePrompt(self):
        name = self.prompt_name_input.text().strip()
        if not name:
            self.status_label.setText("Введите имя шаблона.")
            return
        if self.last_focused_prompt == 2:
            text = self.second_prompt_input.toPlainText()
        else:
            text = self.prompt_input.toPlainText()

        self.prompts_dict[name] = text
        with open(PROMPTS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.prompts_dict, f, ensure_ascii=False, indent=2)
        self.status_label.setText(f"Шаблон '{name}' сохранен.")
        self.loadPromptsToCombo()

    def loadSelectedPrompt(self):
        sel = self.saved_prompts_combo.currentText()
        if sel in self.prompts_dict:
            if self.last_focused_prompt == 2:
                self.second_prompt_input.setPlainText(self.prompts_dict[sel])
            else:
                self.prompt_input.setPlainText(self.prompts_dict[sel])
            self.status_label.setText(f"Шаблон '{sel}' загружен.")
        else:
            self.status_label.setText("Не найден шаблон в базе.")

    def deleteSelectedPrompt(self):
        sel = self.saved_prompts_combo.currentText()
        if sel in self.prompts_dict:
            del self.prompts_dict[sel]
            with open(PROMPTS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.prompts_dict, f, ensure_ascii=False, indent=2)
            self.status_label.setText(f"Шаблон '{sel}' удален.")
            self.loadPromptsToCombo()
        else:
            self.status_label.setText("Не найден шаблон для удаления.")

    def loadVariables(self):
        from pathlib import Path
        if not Path(VARIABLES_FILE).exists():
            self.variables_dict = {}
        else:
            with open(VARIABLES_FILE, "r", encoding="utf-8") as f:
                self.variables_dict = json.load(f)
        self.refreshVariablesTable()

    def refreshVariablesTable(self):
        self.variables_table.setRowCount(0)
        for i, (var_name, col_index) in enumerate(self.variables_dict.items()):
            self.variables_table.insertRow(i)
            it_name = QTableWidgetItem(var_name)
            self.variables_table.setItem(i, 0, it_name)
            it_idx = QTableWidgetItem(str(col_index))
            self.variables_table.setItem(i, 1, it_idx)

            w = QWidget()
            l = QHBoxLayout(w)
            l.setContentsMargins(0,0,0,0)
            b_insert = QPushButton("+")
            b_insert.setToolTip("Вставить переменную")
            b_insert.clicked.connect(lambda _, v=var_name: self.insertVariableInPrompt(v))
            l.addWidget(b_insert)

            b_del = QPushButton("X")
            b_del.setToolTip("Удалить переменную")
            b_del.setStyleSheet("color: red;")
            b_del.clicked.connect(lambda _, v=var_name: self.deleteVariable(v))
            l.addWidget(b_del)

            self.variables_table.setCellWidget(i, 2, w)

        self.variables_table.resizeColumnsToContents()

    def addVariable(self):
        v = self.var_name_input.text().strip()
        c = self.var_col_index.value()
        if not v:
            QMessageBox.warning(self, "Ошибка", "Введите имя переменной")
            return
        if v in self.variables_dict:
            QMessageBox.warning(self, "Ошибка", f"Переменная '{v}' уже есть")
            return

        self.variables_dict[v] = c
        with open(VARIABLES_FILE, "w", encoding="utf-8") as f:
            json.dump(self.variables_dict, f, ensure_ascii=False, indent=2)
        self.var_name_input.clear()
        self.var_col_index.setValue(0)
        self.refreshVariablesTable()
        self.status_label.setText(f"Переменная '{v}' добавлена.")

    def insertVariableInPrompt(self, var_name):
        if self.last_focused_prompt == 2:
            c = self.second_prompt_input.textCursor()
            c.insertText(f" {{{var_name}}}")
            self.second_prompt_input.setTextCursor(c)
        else:
            c = self.prompt_input.textCursor()
            c.insertText(f" {{{var_name}}}")
            self.prompt_input.setTextCursor(c)

    def deleteVariable(self, var_name):
        if var_name in self.variables_dict:
            del self.variables_dict[var_name]
            with open(VARIABLES_FILE, "w", encoding="utf-8") as f:
                json.dump(self.variables_dict, f, ensure_ascii=False, indent=2)
            self.refreshVariablesTable()
            self.status_label.setText(f"Переменная '{var_name}' удалена.")
        else:
            QMessageBox.warning(self, "Ошибка", f"Нет переменной '{var_name}'.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = RewriterApp()
    w.show()
    sys.exit(app.exec_())
