import sys

import pandas as pd
from PyQt5.QtGui import QIcon, QCursor
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QTextEdit, QMessageBox
from PyQt5.QtCore import Qt, QTimer
from openpyxl.styles import Border, Font, Side
from qt_material import apply_stylesheet

import colors
import parser
# import parser_2 as parser


class FileUploader(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_file = None  # Инициализация переменной для хранения выбранного файла
        self.initUI()
        self.setStyleSheet("background-color: white;")
        self.setWindowIcon(QIcon('icon.png'))

    def initUI(self):
        self.setWindowTitle('Ежегодный отчет')

        layout = QVBoxLayout()

        self.label = QLabel('Перетащите файл сюда или нажмите кнопку для выбора файла', self)
        layout.addWidget(self.label)

        self.text_edit = QTextEdit(self)
        self.text_edit.setAcceptDrops(True)
        self.text_edit.dragEnterEvent = self.dragEnterEvent
        self.text_edit.dropEvent = self.dropEvent
        self.text_edit.setStyleSheet("border-color: #66B173;color: black;")
        layout.addWidget(self.text_edit)

        self.notification_label = QLabel('', self)
        self.notification_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.notification_label)

        self.upload_button = QPushButton('Загрузить файл', self)
        self.upload_button.setStyleSheet("background-color: #66B173; color: white; border: none;")
        self.upload_button.setCursor(QCursor(Qt.PointingHandCursor))  # Указатель при наведении
        self.upload_button.clicked.connect(self.loadFile)
        layout.addWidget(self.upload_button)

        self.process_button = QPushButton('Обработать файл', self)
        self.process_button.setStyleSheet("background-color: #66B173; color: white; border: none;")
        self.process_button.setCursor(QCursor(Qt.PointingHandCursor))  # Указатель при наведении
        self.process_button.clicked.connect(self.processFile)
        self.process_button.setEnabled(False)  # Кнопка недоступна изначально
        self.updateProcessButtonStyle()  # Обновляем стиль кнопки
        layout.addWidget(self.process_button)

        self.setLayout(layout)

    def loadFile(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Выберите файл", "", "Все файлы (*)")
        if file_name:
            self.selected_file = file_name
            self.text_edit.setText(file_name.split('/')[-1])  # Отображаем только имя файла
            self.process_button.setEnabled(True)  # Активируем кнопку "Обработать файл"
            self.updateProcessButtonStyle()  # Обновляем стиль кнопки
            self.showNotification("Файл загружен успешно!", success=True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            self.selected_file = event.mimeData().urls()[0].toLocalFile()  # Сохраняем полный путь к файлу
            self.text_edit.setText(self.selected_file.split('/')[-1])  # Отображаем только имя файла
            self.process_button.setEnabled(True)  # Активируем кнопку "Обработать файл"
            self.updateProcessButtonStyle()  # Обновляем стиль кнопки
            self.showNotification("Файл загружен успешно!", success=True)

    def updateProcessButtonStyle(self):
        if self.process_button.isEnabled():
            self.process_button.setStyleSheet("background-color: #66B173; color: white; border: none;")
        else:
            self.process_button.setStyleSheet("background-color: lightgray; color: darkgray; border: none;")  # Серый фон для недоступной кнопки

    def showNotification(self, message, success):
        notification = QWidget(self)
        notification.setWindowTitle("Уведомление")
        notification.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        notification.setAttribute(Qt.WA_TranslucentBackground)

        layout = QVBoxLayout(notification)
        label = QLabel(message)
        label.setStyleSheet("background-color: lightgreen;" if success else "background-color: lightcoral;")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        notification.resize(self.width(), (10 * self.height())//100)  # Установить ширину по ширине родительского окна
        notification.show()

        QTimer.singleShot(2000, notification.close)  # Закрыть уведомление через 2 секунды

    def hideNotification(self):
        self.notification_label.hide()

    def processFile(self):
        sheet = parser.load_file(path=self.selected_file)
        parser.generate_table_ranges(sheet=sheet)
        result = parser.iterate_over_table_ranges(sheet=sheet)
        processed_file_content = parser.create_report(result)

        options = QFileDialog.Options()
        save_file_name, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Текстовые файлы (*.xlsx);;Все файлы (*)", options=options)
        try:
            # # Установка максимальной ширины столбца
            pd.set_option('max_colwidth', 300)
            # Сохранение DataFrame в CSV
            with pd.ExcelWriter(f"{save_file_name}", engine='openpyxl') as writer:
                processed_file_content.to_excel(writer, sheet_name='Report', index=False, header=False)
                worksheet = writer.sheets['Report']
                # Применяем стили к заголовкам
                bold_font = Font(bold=True)
                thin_border = Border(left=Side(style='thin'),
                                     right=Side(style='thin'),
                                     top=Side(style='thin'),
                                     bottom=Side(style='thin'))

                # Применяем цвет к заголовкам (первая строка)
                for col in worksheet.iter_cols(min_row=1, max_row=1):  # Только первая строка
                    for cell in col:
                        colour = colors.header_colours(cell.value)
                        cell.fill = colour
                        cell.font = bold_font  # Применяем жирный шрифт к заголовкам
                        cell.border = thin_border  # Применяем границы к заголовкам

                # Устанавливаем ширину для первых двух столбцов
                for column in worksheet.iter_cols(min_row=1, max_row=1):
                    column_letter = column[0].column_letter
                    if column[0].column <= 2:  # Для первых двух столбцов
                        worksheet.column_dimensions[column_letter].width = 40
                    else:  # Для остальных столбцов
                        worksheet.column_dimensions[column_letter].width = 30

            self.showNotification("Файл сохранен успешно!", success=True)
        except BaseException as e:
            err = e
            self.showNotification("Ошибка при сохранении файла (убедитесь в корректности формата).", success=False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='light_lightgreen.xml')
    ex = FileUploader()
    ex.show()
    sys.exit(app.exec_())
