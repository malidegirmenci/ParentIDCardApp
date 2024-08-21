import openpyxl
from openpyxl.drawing.image import Image

dataFile = 'ParentInfoList.xlsx'
wb = openpyxl.load_workbook(dataFile)
ws_data = wb.active

ws_id_cards = wb.create_sheet(title="Kimlik Kartları")

ws_id_cards.row_dimensions[1].height = 10
ws_id_cards.column_dimensions['A'].width = 2
ws_id_cards.column_dimensions['B'].width = 14
ws_id_cards.column_dimensions['C'].width = 14
ws_id_cards.column_dimensions['D'].width = 14

start_row = 2

for index, row in enumerate(ws_data.iter_rows(min_row=2, values_only=True), start=2):
    tc_id_no, name, surname, classStudent = row

    ws_id_cards.merge_cells(f'B{start_row}:D{start_row+3}')
    ws_id_cards[f'B{start_row+5}'] = 'T.C Kimlik No    :'
    ws_id_cards[f'C{start_row+5}'] = str(tc_id_no)
    ws_id_cards[f'B{start_row+6}'] = 'Adı                   :'
    ws_id_cards[f'C{start_row+6}'] = str(name)
    ws_id_cards[f'B{start_row+7}'] = 'Soyadı              :'
    ws_id_cards[f'C{start_row+7}'] = str(surname)
    ws_id_cards[f'B{start_row+8}'] = 'Sınıfı                :'
    ws_id_cards[f'C{start_row+8}'] = str(classStudent)
    ws_id_cards.merge_cells(f'D{start_row+5}:D{start_row+9}')

    try:
        image_path_top = 'img/BKKimlikUst.png'
        img_top = Image(image_path_top)
        img_top.width = 310
        img_top.height = 80
        ws_id_cards.add_image(img_top, f'B{start_row}')
    except FileNotFoundError:
        print(f"{image_path_top} could not found.")

    try:
        image_path_bottom = 'img/BKKimlikAlt.png'
        img_bottom = Image(image_path_bottom)
        img_bottom.width = 310
        img_bottom.height = 10
        ws_id_cards.add_image(img_bottom, f'B{start_row+10}')
    except FileNotFoundError:
        print(f"{image_path_bottom} could not found.")

    start_row += 15


wb.save('ParentIDCardList.xlsx')
print('Created ID Cards')