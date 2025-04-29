import sys
import requests
import pandas as pd
from datetime import datetime
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget,
                             QTableWidgetItem, QFileDialog, QMessageBox, QStatusBar)
from PyQt6.QtCore import Qt


class HHVacancyParser(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Парсер вакансий hh.ru')
        self.setGeometry(100, 100, 900, 600)

        self.vacancies = []
        self.init_ui()

    def init_ui(self):
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

        # Регион
        params_layout.addWidget(QLabel('Регион:'))
        self.area_combo = QComboBox()
        self.area_combo.addItems(['Москва (1)',
                                  'Санкт-Петербург (2)',
                                  'Россия (113)',
                                  ])
        params_layout.addWidget(self.area_combo)

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

        search_layout.addLayout(params_layout)
        main_layout.addWidget(search_group)

        # Тип занятости
        params_layout.addWidget(QLabel('Занятость'))
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
        self.results_table.setColumnCount(6)
        self.results_table.setHorizontalHeaderLabels(['Название',
                                                      'Компания',
                                                      'Зарплата',
                                                      'Город',
                                                      'Дата публикации',
                                                      'Ссылка',
                                                      ])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        main_layout.addWidget(self.results_table)

        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def search_vacancies(self):
        query = self.query_edit.text().strip()
        if not query:
            QMessageBox.warning(self, 'Ошибка', 'Введите поисковый запрос')
            return

        # Требование запроса API
        headers = {'User-Agent': 'hh.ru-job-parser/1.0 (ivan.karelin.1993@mail.ru)'}

        # URL
        base_url = 'https://api.hh.ru/vacancies'

        # Подготовка параметров
        params = {
            'text': query,
            'area': self.area_combo.currentText().split("(")[1][:-1],
            'per_page': 100,
            'page' : 0
        }

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
        experience_map = {
            'Полная': 'full',
            'Частичная': 'part',
            'Проектная': 'project',
            'Стажировка': 'intern',
            'Волонтёрство': 'volunteer',
        }
        if self.employment_combo.currentText() != 'Любая':
            params['employment'] = experience_map[self.employment_combo.currentText()]

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

        for vacancy in self.vacancies:
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

            # Город
            self.results_table.setItem(row_position, 3,
                                       QTableWidgetItem(vacancy.get('area', {}).get('name', '')))

            # Дата
            pub_date = vacancy.get('published_at', '')
            if pub_date:
                pub_date = datetime.strptime(pub_date, '%Y-%m-%dT%H:%M:%S%z').strftime('%Y-%m-%d')
            self.results_table.setItem(row_position, 4, QTableWidgetItem(pub_date))

            # Ссылка
            link_item = QTableWidgetItem(vacancy.get('alternate_url', ''))
            link_item.setFlags(link_item.flags() ^ Qt.ItemFlag.ItemIsEditable)
            self.results_table.setItem(row_position, 5, link_item)

    def export_to_excel(self):
        """ Функция для сохранения данных в Excel-файл """
        if not self.vacancies:
            QMessageBox.warning(self, 'Предупреждение', 'Нет данных для экспорта')
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self,
            'Сохранить файл',
            f"Результаты поиска {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.xlsx",
            'Excel Files (*.xlsx)'
        )

        if not filepath:
            return

        try:
            data = []
            for vacancy in self.vacancies:
                salary = vacancy.get('salary')
                salary_str = 'Не указана'
                if salary:
                    salary_from = salary.get('from')
                    salary_to = salary.get('to')
                    currency = salary.get('currency', '').upper()
                    salary_str = f"{salary_from or '?'} - {salary_to or '?'} {currency}"

                data.append({
                    'Название': vacancy.get('name'),
                    'Компания': vacancy.get('employer', {}).get('name'),
                    'Зарплата': salary_str,
                    'Город': vacancy.get('area', {}).get('name'),
                    'Опыт': vacancy.get('experience', {}).get('name'),
                    'Тип занятости': vacancy.get('employment', {}).get('name'),
                    'Дата публикации': vacancy.get('published_at'),
                    'Ссылка': vacancy.get('alternate_url')
                })

            df = pd.DataFrame(data)
            if 'Дата публикации' in df.columns:
                df['Дата публикации'] = pd.to_datetime(df['Дата публикации'])

            with pd.ExcelWriter(filepath, engine="xlsxwriter") as file:
                df.to_excel(file, index=False, sheet_name="Вакансии")

                worksheet = file.sheets['Вакансии']
                for i, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(col))
                    worksheet.set_column(i, i, max_len + 2)

            QMessageBox.information(self, 'Успех', f'Файл успешно сохранен:\n{filepath}')
            self.status_bar.showMessage(f'Файл экспортирован: {os.path.basename(filepath)}')

        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Не удалось сохранить файл: {str(e)}')
            self.status_bar.showMessage('Ошибка при экспорте')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = HHVacancyParser()
    window.show()
    sys.exit(app.exec())
