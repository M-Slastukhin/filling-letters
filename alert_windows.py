import PySimpleGUI as sg

def no_landusers_list():
    sg.theme('BluePurple')
    layout = [[sg.Text('Добавьте фаил landusers_list.xlsx ')]]
              #[sg.Button('Выход')]]
    window = sg.Window('Заполнение писем', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Выход':  # закрытие окна
            break
    window.close()
