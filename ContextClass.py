class ContextClass:

    def __init__(self, lander_data):
        self.lander_data = lander_data
        self.context = {}

    def __repr__(self):
        return f'{self.context}'

    __str__ = __repr__

    # Пол адресата
    def get_gender(self, gender):
        #gender = self.lander_data[6]
        if gender[0].lower() == 'м':
            return 'male'
        elif gender[0].lower() == 'ж':
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

    # Фамилия родительный падеж
    @staticmethod
    def _get_surname_dative(surname: str, gender: str) -> str:
        vowels = ['а', 'е', 'ё', 'и', 'й', 'о', 'у', 'э', 'ю', 'я']
        if surname[-2:] == 'ко':
            surname_d = surname
        elif surname[-2:] == 'ия':
            surname_d = surname[:-1] + 'и'
        elif gender == 'female':
            if surname[-3:] == 'ина':
                surname_d = surname[:-1] + 'ой'
            elif surname[-3:] == 'ова':
                surname_d = surname[:-1] + 'ой'
            elif surname[-3:] == 'ева':
                surname_d = surname[:-1] + 'ой'
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

    def update(self, lander_data):
        self.context = {
            'position': lander_data[2],
            'position_d': self._get_position_dative(lander_data[2]),
            'organization': lander_data[1],
            'organization_r': self._get_organization(lander_data[1]),
            'initials': self._get_initials(lander_data[4], lander_data[5]),
            'name': lander_data[4].capitalize(),
            'patronymic': lander_data[5].capitalize(),
            'surname': lander_data[3].capitalize(),
            'surname_d': self._get_surname_dative(lander_data[3], self.get_gender(lander_data[6])),
            'accoct': self._get_accoct(self.get_gender(lander_data[6]))
        }

    def get(self) -> dict:
        return self.context

# if __name__ == "__main__":
