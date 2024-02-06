import openpyxl
from openpyxl.drawing.image import Image

# Создание нового документа Excel
workbook = openpyxl.Workbook()
sheet = workbook.active

# Список названий файлов скачанных картинок
file_names = ['image1.jpg', 'image2.jpg', 'image3.jpg']

# Добавление картинок и номеров этикеток в документ Excel
for i, file_name in enumerate(file_names):
    # Загрузка изображения
    img = Image(file_name)

    # Добавление изображения в ячейку
    cell = sheet.cell(row=i+1, column=1)
    cell.value = i+1
    sheet.add_image(img, 'B{}'.format(i+1))

# Сохранение документа
workbook.save('лист_подбора.xlsx')
