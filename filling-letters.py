import PySimpleGUI as sg
from fl_def import choice_data, main
from fl_doc import mail_save, act_save, plan_save

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
        mail_save(context)
    if event == 'Создать акт':
        act_save(context)
    if event == 'Создать схему':
        plan_save(context)
    if event == sg.WIN_CLOSED or event == 'Выход': # закрытие окна
        break
    #print('You entered ', values[0])


window.close()