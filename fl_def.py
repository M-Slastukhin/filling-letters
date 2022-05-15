import openpyxl
from alert_windows import no_landusers_list

try:
    wb = openpyxl.load_workbook('landusers_list.xlsx')
    sheet = wb.active
    rows = sheet.max_row
    b_rows = 'B' + str(rows)

# Получение перечня оргвнизаций
    choice_data = []
    for column in sheet['B2':str(b_rows)]:
        for cell in column:
            if cell.value is None or cell.value == " ":
               continue
            choice_data.append(cell.value)
except FileNotFoundError:
    no_landusers_list()


# Получение данных по выбранной организации
def lander_info(choice_index=0) -> list:
    lander_data = []
    for row in sheet['A' + str(choice_index+2):'G' + str(choice_index+2)]:
        for cell in row:
            if cell.value is None or cell.value == " ":
                continue
            lander_data.append(cell.value)
    #print(lander_data)
    return lander_data

# получение навания для файла - название организации без ковычек и символов, недопустимых для имени файла
def get_file_name(organization_name):
    file_name = organization_name.replace('?', '')
    file_name = file_name.replace('"', '')
    file_name = file_name.replace('\\', '')
    file_name = file_name.replace('|', ' ')
    file_name = file_name.replace('/', ' ')
    file_name = file_name.replace(':', ' ')
    return file_name