import PySimpleGUI as sg


def no_landusers_list():
    sg.theme('BluePurple')
    layout = [[sg.Text('Отсутствует фаил landusers_list.xlsx')],
              [sg.Button('Выход')]]
    window = sg.Window('Заполнение писем', layout, element_justification='c')
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Выход':  # закрытие окна
            break
    window.close()


def no_mail():
    sg.popup('Отсутствует шаблон письма', 'Добавьте документ "письмо_шаблон.docx" в папку с программой')


def no_act():
    sg.popup('Отсутствует шаблон акта', 'Добавьте документ "акт_шаблон.docx" в папку с программой')
