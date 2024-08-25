import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import Font

dataFile = 'ParentInfoList.xlsx'
wb = openpyxl.load_workbook(dataFile)
ws_data = wb.active
ws_id_cards = wb.create_sheet(title="Kimlik Kartları")

ws_id_cards.page_setup.orientation = ws_id_cards.ORIENTATION_PORTRAIT
ws_id_cards.page_setup.paperSize = ws_id_cards.PAPERSIZE_A4
ws_id_cards.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.75)
arial_font = Font(name='Arial', size=11)

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

#checks text and return reformatted
def formats_text(text):
  if len(text) <= 25:
    return text
  else:
    arr_text = text.split(" ")
    last_index = len(arr_text)-1
    formatted_last_word = arr_text[last_index][0] + '.'
    arr_text.pop(last_index)
    formatted_text = ' '.join(arr_text) + ' ' + formatted_last_word
    return formatted_text

dataFile = 'ParentInfoList.xlsx'
wb = openpyxl.load_workbook(dataFile)
ws_data = wb.active
ws_id_cards = wb.create_sheet(title="Kimlik Kartları")

ws_id_cards.page_setup.orientation = ws_id_cards.ORIENTATION_PORTRAIT
ws_id_cards.page_setup.paperSize = ws_id_cards.PAPERSIZE_A4
ws_id_cards.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
arial_font = Font(name='Arial', size=11)

ws_id_cards.column_dimensions['A'].width = 17
ws_id_cards.column_dimensions['B'].width = 14
ws_id_cards.column_dimensions['C'].width = 14
ws_id_cards.column_dimensions['D'].width = 2
ws_id_cards.column_dimensions['E'].width = 17
ws_id_cards.column_dimensions['F'].width = 14
ws_id_cards.column_dimensions['G'].width = 14

start_row = 1
start_col = 0

for index, row in enumerate(ws_data.iter_rows(min_row=2, values_only=True)):

    parent_name, student_name, level_of_student, class_of_student = row

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
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 5}'] = 'Veli Ad Soyad'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 5}'] = f': {formats_text(parent_name)}'
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 6}'] = 'Öğrenci Ad Soyad'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 6}'] = f': {formats_text(student_name)}'
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 7}'] = 'Kademe'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 7}'] = f': {level_of_student}'
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 8}'] = 'Sınıf'
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 8}'] = f': {class_of_student}'

    ws_id_cards[f'{chr(65 + current_col)}{start_row + 5}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 5}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 6}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 6}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 7}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 7}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col)}{start_row + 8}'].font = arial_font
    ws_id_cards[f'{chr(65 + current_col + 1)}{start_row + 8}'].font = arial_font

    try:
        image_path_bottom = 'img/BKKimlikAlt.png'
        img_bottom = Image(image_path_bottom)
        img_bottom.width = 310
        img_bottom.height = 10
        ws_id_cards.add_image(img_bottom, f'{chr(65 + current_col)}{start_row + 10}')
    except FileNotFoundError:
        print(f"{image_path_bottom} could not be found.")

    if index % 2 == 1:
        start_row += 12

wb.save('ParentIDCardList.xlsx')
print('Created ID Cards')
