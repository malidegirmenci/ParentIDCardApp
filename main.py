import openpyxl
from openpyxl.drawing.image import Image

dataFile = 'ParentInfoList.xlsx'
wb = openpyxl.load_workbook(dataFile)
ws_data = wb.active
ws_id_cards = wb.create_sheet(title="Kimlik Kartları")

ws_id_cards.row_dimensions[1].height = 10
ws_id_cards.column_dimensions['A'].width = 2
ws_id_cards.column_dimensions['B'].width = 16
ws_id_cards.column_dimensions['C'].width = 14
ws_id_cards.column_dimensions['D'].width = 14
ws_id_cards.column_dimensions['E'].width = 2
ws_id_cards.column_dimensions['F'].width = 16
ws_id_cards.column_dimensions['G'].width = 14
ws_id_cards.column_dimensions['H'].width = 14

start_row = 2
start_col = 1

for index, row in enumerate(ws_data.iter_rows(min_row=2, values_only=True)):
    parentName, studentName, levelOfStudent, classOfStudent = row

    col_offset = (index % 2) * 4  # Eğer index çiftse sola yerleşir, tekse sağa yerleşir.
    current_col = start_col + col_offset

    try:
        image_path_top = 'img/BKKimlikUst.png'
        img_top = Image(image_path_top)
        img_top.width = 310
        img_top.height = 80
        ws_id_cards.add_image(img_top, f'{chr(65 + current_col)}{start_row}')
    except FileNotFoundError:
        print(f"{image_path_top} could not be found.")

    ws_id_cards.merge_cells(f'{chr(65 + current_col)}{start_row + 1}:{chr(65 + current_col + 2)}{start_row + 4}')
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 6}'] = 'Veli Ad Soyad       :'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 6}'] = str(parentName)
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 7}'] = 'Öğrenci Ad Soyadı:'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 7}'] = str(studentName)
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 8}'] = 'Kademe               :'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 8}'] = str(levelOfStudent)
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 9}'] = 'Sınıf                     :'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 9}'] = str(classOfStudent)

    try:
        image_path_bottom = 'img/BKKimlikAlt.png'
        img_bottom = Image(image_path_bottom)
        img_bottom.width = 310
        img_bottom.height = 10
        ws_id_cards.add_image(img_bottom, f'{chr(65 + current_col)}{start_row + 11}')
    except FileNotFoundError:
        print(f"{image_path_bottom} could not be found.")

    if index % 2 == 1:  # İki kimlik yan yana yerleştirildikten sonra bir alt satıra geçilir.
        start_row += 15

wb.save('ParentIDCardList.xlsx')
print('Created ID Cards')
