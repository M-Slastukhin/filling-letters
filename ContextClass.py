class ContextClass:

    def __init__(self, lander_data):
        self.lander_data = lander_data
        self.context = {}

    def __repr__(self):
        return f'{self.context}'

    __str__ = __repr__

    # Пол адресата
    @staticmethod
    def _get_gender(gender_ru):
        if gender_ru[0].lower() == 'м':
            return 'male'
        elif gender_ru[0].lower() == 'ж':
            return 'female'
        else:
            return 'male'

    # Определение обращения
    @staticmethod
    def _get_accoct(gender: str) -> str:
        if gender == 'female':
            return 'Уважаемая'
        return 'Уважаемый'

    # Инициалы
    @staticmethod
    def _get_initials(name: str, patronymic: str) -> str:
        return f'{name[0]}.{patronymic[0]}.'

    # Фамилия дательный падеж
    @staticmethod
    def _get_surname_dative(surname: str, gender: str) -> str:
        vowels = ['а', 'е', 'ё', 'и', 'о', 'у', 'э', 'ю', 'я']
        f_endings = ['ина', 'ына', 'ова', 'ева']
        m_endings = ['ый', 'ий', 'ой']
        if surname[-2:] == 'ко':
            return surname.capitalize()
        elif surname[-2:] == 'ия':
            return surname[:-1].capitalize() + 'е'
        elif gender == 'female':
            if any(surname[-3:] == ending for ending in f_endings):
                return surname[:-1].capitalize() + 'ой'
            elif surname[-2:] == 'ая':
                return surname[:-2].capitalize() + 'ой'
            else:
                return surname.capitalize()
        if gender == 'male':
            if surname[-3:] == 'ний':
                return surname[:-2].capitalize() + 'ему'
            elif any(surname[-2:] == ending for ending in m_endings):
                return surname[:-2].capitalize() + 'ому'
            elif surname[-1:] == 'а':
                return surname[:-1].capitalize() + 'е'
            elif any(surname[-1] == v for v in vowels):
                return surname.capitalize()
            elif surname[-1] == 'ь':
                return surname[:-1].capitalize() + 'ю'
            elif surname[-1] == 'й':
                return surname.capitalize()
            elif surname[-2:] == 'ец':
                if surname[-3] == 'л':
                    return surname[:-2].capitalize() + 'ьцу'
                elif any(surname[-3] == v for v in vowels):
                    return surname[:-2].capitalize() + 'йцу'
                else:
                    return surname[:-2].capitalize() + 'цу'
            else:
                return surname.capitalize() + 'у'


# Должность адресата в дательном падеже
    @staticmethod
    def _get_position_dative(position: str):
        split_str = position.split(' ')
        if len(split_str) == 1:
            if position.lower() == 'глава':
                position_d = position[:-1] + 'е'
            elif position[-2:] == 'ль':
                position_d = position[:-1] + 'ю'
            else:
                position_d = position + 'у'
        elif len(split_str) == 2:
            if split_str[0][-2:] == 'ый':
                split_str[0] = split_str[0][:-2] + 'ому'
                split_str[1] = split_str[1] + 'у'
                position_d = f'{split_str[0]} {split_str[1]}'
            else:
                position_d = position
        else:
            position_d = position
        return position_d.capitalize()

    # склонение названия организации
    @staticmethod
    def _get_organization(organization: str):
        split_org = organization.split(' ')
        if split_org[0].lower() == 'администрация':
            split_org[0] = split_org[0][:-1] + 'и'
            split_org[0] = split_org[0].lower()
        elif split_org[0].lower() == 'территориальное':
            split_org[0] = split_org[0][:-1] + 'го'
            split_org[0] = split_org[0].lower()
        if split_org[1].lower() == 'управление':
            split_org[1] = split_org[1][:-1] + 'я'
        else:
            return organization
        org_r = ' '.join(split_org)
        return org_r

    @staticmethod
    def _get_territory(type_of_organization, organization_r):
        if type_of_organization.lower() == 'строительная организация':
            return 'Саратова и Саратовской области'
        elif type_of_organization.lower() == 'землепользователь':
            return organization_r

    # получение навания для файла - название организации без ковычек и символов, недопустимых для имени файла
    @staticmethod
    def _get_file_name(organization):
        file_name = organization.replace('?', '')
        file_name = file_name.replace('"', '')
        file_name = file_name.replace('\\', '')
        file_name = file_name.replace('|', ' ')
        file_name = file_name.replace('/', ' ')
        file_name = file_name.replace(':', ' ')
        return file_name


    def update(self, lander_data):
        self.context = {
            'position': lander_data[3],
            'position_d': self._get_position_dative(lander_data[3]),
            'organization': lander_data[2],
            'organization_r': self._get_organization(lander_data[2]),
            'initials': self._get_initials(lander_data[5], lander_data[6]),
            'name': lander_data[5].capitalize(),
            'patronymic': lander_data[6].capitalize(),
            'surname': lander_data[4].capitalize(),
            'surname_d': self._get_surname_dative(lander_data[4], self._get_gender(lander_data[7])),
            'accoct': self._get_accoct(self._get_gender(lander_data[7])),
            'adress': lander_data[8],
            'tel': lander_data[9],
            'territory': self._get_territory(lander_data[1], self._get_organization(lander_data[2])),
            'file_name': self._get_file_name(lander_data[2])
        }


    def get(self) -> dict:
        return self.context
