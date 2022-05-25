import openpyxl
from alert_windows import no_xlsx

    # Получение перечня оргвнизаций
def excel_read(xlsx_file):
    choice_data = []
    try:
        wb = openpyxl.load_workbook(xlsx_file)
        global sheet
        sheet = wb.active
        for column in sheet['C2':'C' + str(sheet.max_row)]:
            for cell in column:
                if cell.value is None or cell.value == " ":
                    continue
                choice_data.append(cell.value)
        return choice_data
    except openpyxl.utils.exceptions.InvalidFileException:
        no_xlsx()

# Получение данных по выбранной организации
def lander_info(choice_index) -> list:
    lander_data = []
    for row in sheet['A' + str(choice_index+2):'J' + str(choice_index+2)]:
        for cell in row:
            if cell.value is None or cell.value == " ":
                continue
            lander_data.append(cell.value)
    return lander_data
