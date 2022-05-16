import PySimpleGUI as sg
from fl_def import choice_data, lander_info
from fl_doc import mail_save, act_save, plan_save
from alert_windows import user_guide
from ContextClass import ContextClass

lander_data = lander_info()  # полчучение информации из exel по-умолчанию для нулевого землепользователя
info = ContextClass(lander_data)  # создание объекта info (по-умолчанию для нулевого землепользователя)

# интерфейс
sg.theme('BluePurple')   # цвет подложки
# наполнение страницы. текст, выподающий список и кнопки
layout = [  [sg.Text('Выбор землепользователя')],
            [sg.InputCombo(choice_data, key='-combo_choice-', enable_events=True)],
            [sg.Button('Создать письмо'), sg.Button('Создать акт'), sg.Button('Создать схему'), sg.Button('Руководство пользователя'), sg.Button('Выход')]]

# создание окна
window = sg.Window('Заполнение писем', layout)


# Цикл событий для обработки «событий» и получения «значений» входных данных
while True:
    event, values = window.read()
    if event == '-combo_choice-':                                   #обработка события "выбор землепользователя":
        choice_index = choice_data.index(values['-combo_choice-'])  #индекс выбранного землепользователя
        lander_data = lander_info(choice_index)                     #полчучение информации из exel по выбраному индексу
        info.update(lander_data)                                 #заполнение context для объекта info для выбранного землепользователя
    if event == 'Создать письмо':
        mail_save(info.get())
    if event == 'Создать акт':
        act_save(info.get())
    if event == 'Создать схему':
        plan_save(info.get())
    if event == 'Руководство пользователя':
        user_guide()
    if event == sg.WIN_CLOSED or event == 'Выход':  # закрытие окна
        break


window.close()