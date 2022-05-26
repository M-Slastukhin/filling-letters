import PySimpleGUI as sg


def no_docx():
    sg.popup('Укажите шаблон в формате *.docx')


def no_plan():
    sg.popup('Укажите шаблон схемы в формате *.vsdx')


def no_xlsx():
    sg.popup('Укажите фаил с исходными данными в формате *.xlsx')


def bad_xlsx():
    sg.popup('Фаил с исходными данными не заполнен или заполнен неверно')


def user_guide():
    sg.popup('Для использования программы необходимо подговотовить файл в формате *.xlsx '
             'с исходными данными:'
             '\nпорядковый номер, тип (землепользователь/строительная организация),'
             ' название организации, должность, фамилия, имя, отчество, '
             ' пол, адрес, телефон' 
             '\n\nПодстановка данных в письмо,'
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
             '{{tel}} - телефон\n'
             '{{territory}} - для землепользователей название организации в родительном падеже,'
             ' для строительных организаций "Саратова и Саратовской области"'
             )


def successful(file_name):
    sg.popup(f'Создан фаил \"{file_name}\"')
