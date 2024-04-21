import csv
import os
from docxtpl import DocxTemplate

marathon_path = './data_marathon.csv'
if not os.path.exists(marathon_path):
    print("Нет файла с марафоном")
    exit()

with open(marathon_path, 'r', encoding="utf-8") as input_csv_file:
    reader = csv.reader(input_csv_file, delimiter=",")
    all_marathons = list(reader)

current_year = None
template = DocxTemplate("Shablon.docx")

number = 0  # Для отслеживания номера марафона

for marathon in all_marathons:
    number += 1  # Увеличиваем номер марафона
    year = marathon[0]
    city = marathon[5]
    winner = marathon[2]
    winner_name = marathon[1]
    time = marathon[4]

    # Проверяем, начинается ли новый год
    if year != current_year:
        # Если это не первый год, добавляем разрыв страницы
        if current_year is not None:
            template.add_page_break()

        current_year = year

    # Добавляем информацию о марафоне

    context = {
        'number': number,
        'year': year,
        'city': city,
        'men': winner_name,
        'time': time,
    }

    template.render(context)

# Сохраняем результат в одном файле
template.save("Marathon_results.docx")
