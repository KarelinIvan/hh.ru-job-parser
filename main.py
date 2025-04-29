import os
import sys

import pandas as pd
import requests
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget,
                             QTableWidgetItem, QMessageBox, QStatusBar, QCompleter, QFileDialog)
from PyQt6.QtCore import Qt, QStringListModel
from PyQt6.QtGui import QStandardItemModel, QStandardItem


class HHVacancyParser(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Парсер вакансий hh.ru')
        self.setGeometry(100, 100, 1000, 700)

        # Кэш для хранения всех регионов {название: id}
        self.areas_cache = {}

        # Загружаем города при инициализации
        self.load_areas()

        self.vacancies = []
        self.init_ui()

    def load_areas(self):
        """Загружает все регионы из API hh.ru при старте приложения"""
        try:
            response = requests.get('https://api.hh.ru/areas')
            response.raise_for_status()
            areas = response.json()

            # Рекурсивная функция для обработки всех уровней вложенности
            def process_area(area):
                self.areas_cache[area['name'].lower()] = area['id']
                for sub_area in area.get('areas', []):
                    process_area(sub_area)

            for area in areas:
                process_area(area)

            print(f"Загружено {len(self.areas_cache)} регионов")  # Для отладки
        except Exception as e:
            print(f"Ошибка при загрузке регионов: {e}")
            # Запасной вариант - основные города
            default_areas = {
                'москва': 1,
                'санкт-петербург': 2,
                'новосибирск': 4,
                'екатеринбург': 3,
                'казань': 88,
                'нижний новгород': 66,
                'челябинск': 104,
                'самара': 78,
                'омск': 68,
                'ростов-на-дону': 76,
                'уфа': 99,
                'красноярск': 54,
                'пермь': 72,
                'воронеж': 26,
                'волгоград': 24,
                'архангельск': 14,
                'северодвинск': 1017,
                'ярославль': 112,
                'иваново': 32
            }
            self.areas_cache.update(default_areas)

    def init_ui(self):
        """ Функция для вывода информации в приложение и фильтры """
        # Главный виджет и layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Панель параметров поиска
        search_group = QWidget()
        search_layout = QVBoxLayout(search_group)

        # Поисковый запрос
        query_layout = QHBoxLayout()
        query_layout.addWidget(QLabel('Поисковый запрос:'))
        self.query_edit = QLineEdit()
        self.query_edit.setPlaceholderText('Например: Python разработчик')
        query_layout.addWidget(self.query_edit)
        search_layout.addLayout(query_layout)

        # Дополнительные параметры
        params_layout = QHBoxLayout()

        # Поле для ввода города
        params_layout.addWidget(QLabel('Город:'))
        self.city_edit = QLineEdit()
        self.city_edit.setPlaceholderText("Например: Москва")

        # Настройка автодополнения для городов
        completer_model = QStringListModel(list(self.areas_cache.keys()))
        completer = QCompleter()
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setModel(completer_model)
        self.city_edit.setCompleter(completer)

        params_layout.addWidget(self.city_edit)

        # Зарплата
        params_layout.addWidget(QLabel('Зарплата от:'))
        self.salary_edit = QLineEdit()
        self.salary_edit.setFixedWidth(100)
        params_layout.addWidget(self.salary_edit)

        # Опыт работы
        params_layout.addWidget(QLabel('Опыт:'))
        self.experience_combo = QComboBox()
        self.experience_combo.addItems(['Любой',
                                        'Нет опыта',
                                        'От 1 года',
                                        'От 3 лет',
                                        'Более 6 лет',
                                        ])
        params_layout.addWidget(self.experience_combo)

        # Тип занятости
        params_layout.addWidget(QLabel('Занятость:'))
        self.employment_combo = QComboBox()
        self.employment_combo.addItems([
            'Любая',
            'Полная',
            'Частичная',
            'Проектная',
            'Стажировка',
            'Волонтёрство',
        ])
        params_layout.addWidget(self.employment_combo)

        # Форма работы
        params_layout.addWidget(QLabel('Формат работы'))
        self.schedule_combo = QComboBox()
        self.schedule_combo.addItems([
            'Любая',
            'Полный день',
            'Сменный график',
            'Гибкий график',
            'Удаленная работа',
            'Вахтовый метод',
        ])
        params_layout.addWidget(self.schedule_combo)

        search_layout.addLayout(params_layout)
        main_layout.addWidget(search_group)

        # Кнопки управления
        button_layout = QHBoxLayout()
        self.search_btn = QPushButton('Найти вакансии')
        self.search_btn.clicked.connect(self.search_vacancies)
        button_layout.addWidget(self.search_btn)

        self.export_btn = QPushButton('Экспорт в Excel')
        self.export_btn.clicked.connect(self.export_to_excel)
        self.export_btn.setEnabled(False)
        button_layout.addWidget(self.export_btn)

        main_layout.addLayout(button_layout)

        # Таблица результатов
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(8)
        self.results_table.setHorizontalHeaderLabels(['Название',
                                                      'Компания',
                                                      'Зарплата',
                                                      'Тип занятости',
                                                      'Форма работы',
                                                      'Город',
                                                      'Дата публикации',
                                                      'Ссылка',
                                                      ])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        main_layout.addWidget(self.results_table)

        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def get_area_id(self, city_name):
        """Возвращает ID региона по названию города"""
        return self.areas_cache.get(city_name.strip().lower())

    def search_vacancies(self):
        """ Функция для получения вакансий через API hh.ru """
        query = self.query_edit.text().strip()
        if not query:
            QMessageBox.warning(self, 'Ошибка', 'Введите поисковый запрос')
            return

        # Требование запроса API
        headers = {'User-Agent': 'hh.ru-job-parser/1.0 (ivan.karelin.1993@mail.ru)'}

        # URL
        base_url = 'https://api.hh.ru/vacancies'

        # Обработка города
        city_name = self.city_edit.text().strip()
        area_id = None

        if city_name:
            self.status_bar.showMessage(f'Поиск вакансий в городе {city_name}...')
            QApplication.processEvents()

            area_id = self.get_area_id(city_name)
            if not area_id:
                QMessageBox.warning(self, 'Ошибка', f'Город "{city_name}" не найден. Проверьте написание.')
                self.status_bar.clearMessage()
                return

        params = {
            'text': query,
            'per_page': 100,
            'page': 0
        }

        if area_id:
            params['area'] = area_id

        # Зарплата
        if self.salary_edit.text():
            try:
                params['salary'] = int(self.salary_edit.text())
            except ValueError:
                QMessageBox.warning(self, 'Ошибка', 'Некорректное значение зарплаты')
                return

        # Опыт работы
        experience_map = {
            'Нет опыта': 'noExperience',
            'От 1 года': 'between1And3',
            'От 3 лет': 'between3And6',
            'Более 6 лет': 'moreThan6',
        }
        if self.experience_combo.currentText() != 'Любой':
            params['experience'] = experience_map[self.experience_combo.currentText()]

        # Тип занятости
        employment_map = {
            'Полная': 'full',
            'Частичная': 'part',
            'Проектная': 'project',
            'Стажировка': 'probation',
            'Волонтёрство': 'volunteer',
        }
        if self.employment_combo.currentText() != 'Любая':
            params['employment'] = employment_map[self.employment_combo.currentText()]

        # Форма работы
        schedule_map = {
            'Полный день': 'fullDay',
            'Сменный график': 'shift',
            'Гибкий график': 'flexible',
            'Удаленная работа': 'remote',
            'Вахтовый метод': 'flyInFlyOut',
        }
        if self.schedule_combo.currentText() != 'Любая':
            params['schedule'] = schedule_map[self.schedule_combo.currentText()]

        self.status_bar.showMessage('Идет поиск вакансий...')
        QApplication.processEvents()

        try:
            response = requests.get(base_url, params=params, headers=headers)
            response.raise_for_status()
            data = response.json()

            self.vacancies = data.get('items', [])
            self.display_results()

            self.status_bar.showMessage(f'Найдено вакансий: {len(self.vacancies)}')
            self.export_btn.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось получить данные: {str(e)}')
            self.status_bar.showMessage('Ошибка при получении данных')

    def display_results(self):
        """ Функуия для отображения данных о вакансиях в таблице в графическом интерфейсе приложения """
        self.results_table.setRowCount(0)

        # Сортируем вакансии по дате (новые сначала)
        sorted_vacancies = sorted(
            self.vacancies,
            key=lambda x: datetime.strptime(x.get('published_at', ''), '%Y-%m-%dT%H:%M:%S%z'),
            reverse=True
        )

        for vacancy in sorted_vacancies:
            row_position = self.results_table.rowCount()
            self.results_table.insertRow(row_position)

            # Название
            self.results_table.setItem(row_position, 0,
                                       QTableWidgetItem(vacancy.get('name', '')))

            # Компания
            self.results_table.setItem(row_position, 1,
                                       QTableWidgetItem(vacancy.get('employer', {}).get('name', '')))

            # Зарплата
            salary = vacancy.get('salary')
            salary_str = 'Не указана'
            if salary:
                salary_from = salary.get('from')
                salary_to = salary.get('to')
                currency = salary.get('currency', '').upper()
                salary_str = f"{salary_from or '0'} - {salary_to or 'Не указано'} {currency}"
            self.results_table.setItem(row_position, 2, QTableWidgetItem(salary_str))

            # Тип занятости
            employment = vacancy.get('employment', {}).get('name', 'Не указан')
            self.results_table.setItem(row_position, 3, QTableWidgetItem(employment))

            # Форма работы
            schedule = vacancy.get('schedule', {}).get('name', 'Не указана')
            self.results_table.setItem(row_position, 4, QTableWidgetItem(schedule))

            # Город
            self.results_table.setItem(row_position, 5,
                                       QTableWidgetItem(vacancy.get('area', {}).get('name', '')))

            # Дата
            pub_date = vacancy.get('published_at', '')
            if pub_date:
                pub_date = datetime.strptime(pub_date, '%Y-%m-%dT%H:%M:%S%z').strftime('%Y-%m-%d')
            self.results_table.setItem(row_position, 6, QTableWidgetItem(pub_date))

            # Ссылка
            link_item = QTableWidgetItem(vacancy.get('alternate_url', ''))
            link_item.setFlags(link_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.results_table.setItem(row_position, 7, link_item)

    def export_to_excel(self):
        """Функция для сохранения данных в Excel-файл"""
        if not self.vacancies:
            QMessageBox.warning(self, 'Предупреждение', 'Нет данных для экспорта')
            return

        # Предлагаем имя файла по умолчанию
        default_filename = f"Результаты поиска {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        filepath, _ = QFileDialog.getSaveFileName(
            self,
            'Сохранить файл',
            default_filename,
            'Excel Files (*.xlsx)'
        )

        if not filepath:
            return  # Пользователь отменил сохранение

        try:
            data = []
            for vacancy in self.vacancies:
                # Обработка зарплаты
                salary = vacancy.get('salary')
                salary_str = 'Не указана'
                if salary:
                    salary_from = salary.get('from', '?')
                    salary_to = salary.get('to', '?')
                    currency = salary.get('currency', '').upper()
                    salary_str = f"{salary_from} - {salary_to} {currency}".strip()

                # Обработка даты (удаление часового пояса)
                pub_date = vacancy.get('published_at')
                if pub_date:
                    try:
                        pub_date = pd.to_datetime(pub_date).tz_localize(None)
                    except:
                        pub_date = None

                data.append({
                    'Название': vacancy.get('name', ''),
                    'Компания': vacancy.get('employer', {}).get('name', ''),
                    'Зарплата': salary_str,
                    'Город': vacancy.get('area', {}).get('name', ''),
                    'Опыт': vacancy.get('experience', {}).get('name', ''),
                    'Тип занятости': vacancy.get('employment', {}).get('name', ''),
                    'Формат работы': vacancy.get('schedule', {}).get('name', ''),
                    'Дата публикации': pub_date,
                    'Ссылка': vacancy.get('alternate_url', '')
                })

            # Создаем DataFrame
            df = pd.DataFrame(data)

            # Убедимся, что даты правильно обработаны
            if 'Дата публикации' in df.columns:
                df['Дата публикации'] = pd.to_datetime(df['Дата публикации'], errors='coerce')

            # Сохраняем в Excel
            with pd.ExcelWriter(filepath, engine='xlsxwriter', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
                df.to_excel(writer, index=False, sheet_name='Вакансии')

                # Настраиваем ширину столбцов
                worksheet = writer.sheets['Вакансии']
                for i, col in enumerate(df.columns):
                    # Определяем максимальную длину содержимого
                    max_len = max(
                        df[col].astype(str).apply(len).max(),  # Макс. длина данных
                        len(str(col))  # Длина заголовка
                    )
                    worksheet.set_column(i, i, min(max_len + 2, 50))  # Ограничиваем максимальную ширину

            QMessageBox.information(self, 'Успех', f'Файл успешно сохранен:\n{filepath}')
            self.status_bar.showMessage(f'Файл экспортирован: {os.path.basename(filepath)}')

        except PermissionError:
            QMessageBox.critical(self, 'Ошибка', 'Нет прав для записи в указанное место')
            self.status_bar.showMessage('Ошибка: нет прав для записи')
        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось сохранить файл:\n{str(e)}')
            self.status_bar.showMessage('Ошибка при экспорте')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = HHVacancyParser()
    window.show()
    sys.exit(app.exec())
