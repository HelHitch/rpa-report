# #!/usr/bin/env python
# # coding: utf-8
#
# # In[115]:
#
#
# import os
# import shutil
# import pandas as pd
# import openpyxl
# from openpyxl.styles import PatternFill, Border, Side, Alignment
#
# # Путь к исходному файлу
# source = 'inputFile.xlsx'
# # Путь, куда нужно скопировать файл
# destination = 'copyFile.xlsx'
#
# # Проверяем, существует ли файл назначения и удаляем его, если да
# if os.path.exists(destination):
#     os.remove(destination)
#     print("Старый файл удален")
#
# # Копируем новый файл
# shutil.copyfile(source, destination)
# print("Файл успешно скопирован")
#
# # Путь к копии файла
# file_path = destination  # Используем уже скопированный файл
#
# # Открываем файл с помощью openpyxl
# wb = openpyxl.load_workbook(file_path)
# ws = wb.active
#
# # Диапазоны для проверки
# start_row, end_row = 14, 84
# start_column, end_column = 'B', 'QY'
#
# # Цвета для поиска (hex-коды)
# colors_to_find = {
#     "yellow": "FFFFFF00",  # Желтый
#     "blue": "FFCCCCFF",    # Голубой
#     "orange": "FFFFC000"   # Оранжевый
# }
#
# # Список для хранения результатов
# results = []
#
# # Проходим по строкам от 14 до 84
# for row in range(start_row, end_row + 1):
#     # Счетчики для каждого цвета
#     yellow_cells, blue_cells, orange_cells = 0, 0, 0
#
#     # Считываем ФИО из столбца B
#     fio = ws[f'B{row}'].value
#
#     # Проходим по столбцам от B до QY в текущей строке
#     for col in range(openpyxl.utils.column_index_from_string(start_column), openpyxl.utils.column_index_from_string(end_column) + 1):
#         cell = ws.cell(row=row, column=col)
#         cell_color = cell.fill.start_color.rgb if cell.fill and cell.fill.start_color else None
#
#         # Проверяем цвет ячейки
#         if cell_color == colors_to_find["yellow"]:
#             yellow_cells += 1
#         elif cell_color == colors_to_find["blue"]:
#             blue_cells += 1
#         elif cell_color == colors_to_find["orange"]:
#             orange_cells += 1
#
#     # Добавляем результаты в список
#     results.append({
#         'ФИО': fio,
#         'Row': row,
#         'Yellow': yellow_cells,
#         'Blue': blue_cells,
#         'Orange': orange_cells
#     })
#
# # Преобразуем список в DataFrame
# results_df = pd.DataFrame(results)
#
# # Записываем результаты в новый Excel файл с помощью pandas
# output_file_path = 'copyFile_results.xlsx'
# results_df.to_excel(output_file_path, index=False, sheet_name='Results')
#
# # Открываем новый файл для записи заливки
# output_wb = openpyxl.load_workbook(output_file_path)
# output_ws = output_wb.active
#
# # Задаем цвета для заголовков
# header_fill_colors = {
#     'Yellow': 'FFFFFF00',  # Желтый
#     'Blue': 'FFCCCCFF',    # Голубой
#     'Orange': 'FFFFC000'   # Оранжевый
# }
#
# # Заливаем заголовки цветами
# header_row = output_ws[1]  # Получаем первую строку (заголовки)
# for cell in header_row:
#     if cell.value == 'Yellow':
#         cell.fill = PatternFill(start_color=header_fill_colors['Yellow'], end_color=header_fill_colors['Yellow'], fill_type='solid')
#     elif cell.value == 'Blue':
#         cell.fill = PatternFill(start_color=header_fill_colors['Blue'], end_color=header_fill_colors['Blue'], fill_type='solid')
#     elif cell.value == 'Orange':
#         cell.fill = PatternFill(start_color=header_fill_colors['Orange'], end_color=header_fill_colors['Orange'], fill_type='solid')
#
# # Заливаем ячейки в новом файле и добавляем обводку
# thin_border = Border(left=Side(style='thin'),
#                      right=Side(style='thin'),
#                      top=Side(style='thin'),
#                      bottom=Side(style='thin'))
#
# for index, row in results_df.iterrows():
#     row_num = int(row['Row'])  # Номер строки
#     yellow_count = row['Yellow']
#     blue_count = row['Blue']
#     orange_count = row['Orange']
#
#     # Заливаем ячейки желтым, голубым и оранжевым цветом в новой таблице и добавляем обводку
#     if yellow_count > 0:
#         output_ws[f'QZ{row_num}'] = yellow_count  # Записываем количество
#         output_ws[f'QZ{row_num}'].fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
#         output_ws[f'QZ{row_num}'].border = thin_border  # Добавляем обводку
#         output_ws[f'QZ{row_num}'].alignment = Alignment(horizontal='center')  # Центрируем текст
#     if blue_count > 0:
#         output_ws[f'RA{row_num}'] = blue_count  # Записываем количество
#         output_ws[f'RA{row_num}'].fill = PatternFill(start_color='FFCCCCFF', end_color='FFCCCCFF', fill_type='solid')
#         output_ws[f'RA{row_num}'].border = thin_border  # Добавляем обводку
#         output_ws[f'RA{row_num}'].alignment = Alignment(horizontal='center')  # Центрируем текст
#     if orange_count > 0:
#         output_ws[f'RB{row_num}'] = orange_count  # Записываем количество
#         output_ws[f'RB{row_num}'].fill = PatternFill(start_color='FFFFC000', end_color='FFFFC000', fill_type='solid')
#         output_ws[f'RB{row_num}'].border = thin_border  # Добавляем обводку
#         output_ws[f'RB{row_num}'].alignment = Alignment(horizontal='center')  # Центрируем текст
#
# # Записываем ФИО в новый столбец и добавляем обводку
# for index, row in results_df.iterrows():
#     output_ws[f'A{index + 2}'] = row['ФИО']  # Записываем ФИО в первую колонку, начиная со второй строки
#     output_ws[f'A{index + 2}'].border = thin_border  # Добавляем обводку
#     output_ws[f'A{index + 2}'].alignment = Alignment(horizontal='center')  # Центрируем текст
#
# # Обводка для других заполненных ячеек
# for row in range(2, len(results_df) + 2):  # Начинаем с 2, так как первая строка - заголовки
#     for col in ['A', 'B', 'C', 'D', 'E']:  # Столбцы, к которым добавляем обводку
#         output_ws[f'{col}{row}'].border = thin_border  # Добавляем обводку ко всем заполненным ячейкам
#
# # Сохраняем изменения в новом файле
# output_wb.save(output_file_path)
# print(f"Результаты записаны и сохранены в файл: {output_file_path}")
#
# # Путь к файлам
# source_file_path = output_file_path
# destination_file_path = destination
#
# # Загружаем оба файла
# source_wb = openpyxl.load_workbook(source_file_path)
# destination_wb = openpyxl.load_workbook(destination_file_path)
#
# # Выбираем активные листы
# source_ws = source_wb.active
# destination_ws = destination_wb.active
#
# # Начинаем вставку данных с QZ13
# start_row = 13
# start_col = openpyxl.utils.column_index_from_string('QZ')
#
# # Копируем данные и стили из source в destination, пропуская вторую колонку
# for source_row in source_ws.iter_rows(min_row=1, max_row=source_ws.max_row, min_col=1, max_col=source_ws.max_column):
#     for cell in source_row:
#         # Пропускаем вторую колонку (индекс 1)
#         if cell.column == 2:  # Если это вторая колонка, пропускаем
#             continue
#
#         # Определяем новую ячейку в destination
#         new_row = start_row + cell.row - 1  # Увеличиваем номер строки на 12 (так как начинаем с QZ13)
#         new_col = start_col + cell.column - 1  # Смещение по столбцам
#         new_cell = destination_ws.cell(row=new_row, column=new_col, value=cell.value)
#
#         # Копируем стиль из source в destination
#         if cell.has_style:
#             new_cell._style = cell._style
#
# # Сохраняем изменения в целевом файле
# destination_wb.save(destination_file_path)
#
# print(f"Данные и стили из {source_file_path} успешно объединены с {destination_file_path}, начиная с QZ13 (вторая колонка пропущена).")
#
#
# # In[ ]:
#
#
#
#
