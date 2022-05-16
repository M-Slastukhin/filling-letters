import PySimpleGUI as sg


def no_landusers_list():
    sg.theme('BluePurple')
    layout = [[sg.Text('Отсутствует фаил landusers_list.xlsx')],
              [sg.Button('Руководство пользователя'), sg.Button('Выход')]]
    window = sg.Window('Заполнение писем', layout, element_justification='c')
    while True:
        event, values = window.read()
        if event == 'Руководство пользователя':
            user_guide()
        if event == sg.WIN_CLOSED or event == 'Выход':  # закрытие окна
            break
    window.close()


def no_mail():
    sg.popup('Отсутствует шаблон письма', 'Добавьте документ "письмо_шаблон.docx" в папку с программой')


def no_act():
    sg.popup('Отсутствует шаблон акта', 'Добавьте документ "акт_шаблон.docx" в папку с программой')


def no_plan():
    sg.popup('Отсутствует шаблон схемы', 'Добавьте документ "схема_шаблон.vsdx" в папку с программой')


def user_guide():
    sg.popup('Для использования программы необходимо разместить подговотовленные файлы '
             'landusers_list.xlsx, письмо_шаблон.docx, акт_шаблон.docx, схема_шаблон.vsdx '
             'в папке с программой.\n\nФаил landusers_list должен содержать исходные данные по '
             'категориям:\n  название организации (organization),\n  должность (position),\n  фамилия (surname),\n '
             ' имя (name),\n  отчество (patronymic),\n  пол (gender),\n  адрес (adress),\n  '
             'тип(землепользователь/строительная организация) (type)\n\nПодстановка данных в письмо,'
             ' акт или схему осуществляется заменой ключевых слов.\n\nДоступные ключевые слова:\n'
             '{{organization}} - название организации\n'
             '{{organization_r}} - название организации в родительном падеже\n'
             '{{position}} - должность\n'
             '{{position_d}} - должность в дательном падеже\n'
             '{{name}} - имя\n'
             '{{patronymic}} - отчество\n'
             '{{surname}} - фамилия\n'
             '{{surname_d}} - фамилия в дательном падеже\n'
             '{{initials}} - инициалы\n'
             '{{accoct}} - обращение (Уважаемый/Уважаемая)\n'
             '{{adress}} - адрес\n'
             '{{territory}} - для землепользователей название организации в родительном падеже,'
             ' для строительных организаций "Саратова и Саратовской области"'
             )