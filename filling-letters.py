import PySimpleGUI as sg
from alert_windows import user_guide, no_xlsx, bad_xlsx
from fl_def import excel_read, lander_info
from fl_doc import docx_save, vsdx_save
from ContextClass import ContextClass

info = ContextClass([])  # создание объекта info (по-умолчанию для нулевого землепользователя)
choice_data = []
# интерфейс
#sg.theme_previewer() возможные вариаты оформления
sg.theme('DefaultNoMoreNagging')   # цвет подложки
# наполнение страницы
layout = [  [sg.In(key='-sample_mail-'), sg.FileBrowse('Выбор шаблона письма', size = (20))],
            [sg.In(key='-sample_act-'), sg.FileBrowse('Выбор шаблона акта', size = (20))],
            [sg.In(key='-sample_plan-'), sg.FileBrowse('Выбор шаблона схемы ', size = (20))],
            [sg.In(key='-landusers_list-', enable_events=True), sg.FileBrowse('Исходные данные', size = (20))],
            [sg.Text('Выбор землепользователя')],
            [sg.InputCombo(choice_data, key='-combo_choice-', enable_events=True, size = (75))],
            [sg.Button('Создать письмо'), sg.Button('Создать акт'), sg.Button('Создать схему'), sg.Button('Руководство пользователя'), sg.Button('Выход')]]

# создание окна
window = sg.Window('Заполнение писем', layout)


# Цикл событий для обработки «событий» и получения «значений» входных данных
while True:
    event, values = window.read()
    if event == '-landusers_list-':
        try:
            choice_data = excel_read(values['-landusers_list-'])
            window['-combo_choice-'].update(values=choice_data)
            info.update(lander_info(0))
        except FileNotFoundError:
            no_xlsx()
        except NameError:
            None
        except IndexError:
            bad_xlsx()
    if event == '-combo_choice-':                                   #обработка события "выбор землепользователя":
        choice_index = choice_data.index(values['-combo_choice-'])  #индекс выбранного землепользователя
        lander_data = lander_info(choice_index)                     #полчучение информации из exel по выбраному индексу
        try:
            info.update(lander_data)                                 #заполнение context для объекта info для выбранного землепользователя
        except IndexError:
            bad_xlsx()
    if event == 'Создать письмо':
        docx_save(info.get(), values['-sample_mail-'], 'письмо')
    if event == 'Создать акт':
        docx_save(info.get(), values['-sample_act-'], 'акт')
    if event == 'Создать схему':
        vsdx_save(info.get(), values['-sample_plan-'])
    if event == 'Руководство пользователя':
        user_guide()
    if event == sg.WIN_CLOSED or event == 'Выход':  # закрытие окна
        break


window.close()
