from docxtpl import DocxTemplate
import openpyxl
import PySimpleGUI as sg


doc_mail = DocxTemplate('письмо_шаблон.docx')
doc_act = DocxTemplate('акт_шаблон.docx')
doc_plan = DocxTemplate('схема_шаблон.docx')
choice_data = []
context = {}
wb = openpyxl.load_workbook('landusers_list.xlsx')
sheet = wb.active
rows = sheet.max_row
b_rows = 'B'+ str(rows)

# Получение перечня оргвнизаций
for column in sheet['B2':str(b_rows)]:
    for cell in column:
        if cell.value is None or cell.value == " ":
            continue
        choice_data.append(cell.value)

# получение данных по выбранной организации
def main(choice):
    landusers_data = []
    choice= choice+2
    A = 'A'+str(choice)
    G = 'G'+str(choice)
    for row in sheet[str(A):str(G)]:
        for cell in row:
            if cell.value is None or cell.value == " ":
                continue
            landusers_data.append(cell.value)
    gender = get_gender(landusers_data[6].lower())

    context = {
        'position': landusers_data[2],
        'position_d': get_position_dative(landusers_data[2]),
        'organization': landusers_data[1],
        'organization_r': get_organization(landusers_data[1]),
        'initials': get_initials(landusers_data[4], landusers_data[5]),
        'name': landusers_data[4].capitalize(),
        'patronymic': landusers_data[5].capitalize(),
        'surname': landusers_data[3].capitalize(),
        'surname_d': get_surname_dative(landusers_data[3], gender),
        'accoct': get_accoct(gender)
    }

    return context

# Пол адресата
def get_gender(gender_form_wb):
        if gender_form_wb[0] == 'м':
            return 'male'
        elif gender_form_wb[0] == 'ж':
            return 'female'
        else:
            return 'male'

# Определение обращения
def get_accoct(gender):
    if gender == 'female':
        return 'Уважаемая'
    return 'Уважаемый'

# Инициалы
def get_initials(name, patronymic):
    return f'{name[0]}.{patronymic[0]}.'

# Фамилия родительный падеж
def get_surname_dative(surname, gender):
    vowels = ['а', 'е', 'ё', 'и', 'й', 'о', 'у', 'э', 'ю', 'я']
    if surname[-2:] == 'ко':
        surname_d = surname
    elif surname[-2:] == 'ия':
        surname_d = surname[:-1]+'и'
    elif gender == 'female':
        if surname[-3:] == 'ина':
            surname_d = surname[:-1]+'ой'
        elif surname[-3:] == 'ова':
            surname_d = surname[:-1]+'ой'
        elif surname[-3:] == 'ева':
            surname_d = surname[:-1]+'ой'
        else:
            surname_d = surname
    elif gender == 'male':
        for i in vowels:
            if surname[-1] == i:
                surname_d = surname
                break
            elif surname.lower() == 'коломиец':
                surname_d = 'коломийцу'
                break
        else:
            surname_d = f'{surname}у'
    return surname_d.capitalize()

# Должность адресата в родительном падеже
def get_position_dative(position):
    split_str = position.split(' ')
    if len(split_str) == 1:
        n = len(position)
        if position.lower() == 'глава':
            position_d = position[0:-1]+'е'
        elif position[-2:n] =='ль':
            position_d = position[0:-1]+'ю'
        else:
            position_d = position+'у'
    elif len(split_str) == 2:
        n = len(split_str[0])
        if split_str[0][-2:n] =='ый':
            split_str[0] = split_str[0][0:-2]+'ому'
            split_str[1] = split_str[1]+'у'
            position_d = f'{split_str[0]} {split_str[1]}'
        else:
            position_d = position
    else:
        position_d = position
    return position_d.capitalize()

# склонение названия организации
def get_organization(organization):
    split_org = organization.split(' ')
    n = len(split_org)
    org_r = ''
    if split_org[0].lower() == 'администрация':
        split_org[0] = split_org[0][0:-1]+'и'
        split_org[0] = split_org[0].lower()
        for i in range (0, n):
            org_r = org_r+' '+split_org[i]
    elif split_org[0].lower() == 'территориальное':
        split_org[0] = split_org[0][0:-1] + 'го'
        split_org[0] = split_org[0].lower()
        if split_org[1].lower() == 'управление':
            split_org[1] = split_org[1][0:-1] + 'я'
        for i in range(0, n):
            org_r = org_r + ' ' + split_org[i]
    else:
        org_r = organization
    return org_r

# получение навания для файла - название организации без ковычек и символов, недопустимых для имени файла
def get_file_name(organization_name):
    file_name = organization_name.replace('?', '')
    file_name = file_name.replace('"', '')
    file_name = file_name.replace('\\', '')
    file_name = file_name.replace('|', ' ')
    file_name = file_name.replace('/', ' ')
    file_name = file_name.replace(':', ' ')
    return file_name

# создание письма
def mail_save():
    doc_mail.render(context)
    file_name = get_file_name(context['organization'])
    doc_mail.save(f'письмо {file_name}.docx')

# создание акта
def act_save():
    doc_act.render(context)
    file_name = get_file_name(context['organization'])
    doc_act.save(f'акт {file_name}.docx')

# создание схемы
def plan_save():
    doc_plan.render(context)
    file_name = get_file_name(context['organization'])
    doc_plan.save(f'схема {file_name}.docx')

# интерфейс
sg.theme('BluePurple')   # цвет подложки
# наполнение страницы.
layout = [  [sg.Text('Выбор землепользователя')],
            [sg.InputCombo(choice_data, key='-combo_choice-', enable_events=True)],
            [sg.Button('Создать письмо'), sg.Button('Создать акт'), sg.Button('Создать схему'), sg.Button('Выход')] ]

# создание окна
window = sg.Window('Заполнение писем', layout)
# Цикл событий для обработки «событий» и получения «значений» входных данных
while True:
    event, values = window.read()
    if event == '-combo_choice-':
        context = main(choice_data.index(values['-combo_choice-']))
        print(choice_data.index(values['-combo_choice-']))
    if event == 'Создать письмо':
        mail_save()
    if event == 'Создать акт':
        act_save()
    if event == 'Создать схему':
        plan_save()
    if event == sg.WIN_CLOSED or event == 'Выход': # закрытие окна
        break
    #print('You entered ', values[0])


window.close()