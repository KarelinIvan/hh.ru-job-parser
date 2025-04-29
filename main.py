from config import get_hh_vacancies
from view import get_user_params, parse_vacancy_date, save_to_excel


def main():
    """ Функция для запуска парсера """
    # Получаем параметры поиска от пользователя
    params = get_user_params()
    # Получаем вакансии
    vacancies = get_hh_vacancies(params)

    if not vacancies:
        print('Не найдено подходящих вакансий')
        return

    # Обрабатывае данные
    parsed_data = [parse_vacancy_date(i) for i in vacancies]

    # Сохраняем в Excel-фаил
    save_to_excel(parsed_data)

if __name__ == '__main__':
    main()
