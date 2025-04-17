from kivy.app import App

from . import DocAnalyzer
from . import DocxAnalyzer


class Checker:
    last_file = None
    result = None
    # Методы проверки для Влада. Пока как заглушки

    def update_result(self):
        selected = App.get_running_app().selected_file
        if selected and self.last_file != selected:
            self.last_file = selected
            if selected.endswith(".doc"):
                self.result = DocAnalyzer.analyze(selected)
            elif selected.endswith(".docx"):
                self.result = DocxAnalyzer.main(selected)
            else:
                self.result = None

    # Для всего (итоговая оценка отчета. Возврат только True или False)

    def query(self, key):
        self.update_result()
        if self.result is None:
            return None
        if key not in self.result:
            return None
        return self.result[key]

    def run_all_checks(self):
        selected = App.get_running_app().selected_file
        if not selected:
            return False
        if selected.endswith(".doc"):
            self.result = DocAnalyzer.analyze(selected)
        elif selected.endswith(".docx"):
            self.result = DocxAnalyzer.main(selected)
        else:
            self.result = None
        if self.result is None:
            return False
        return all([x for x in self.result.values()])

    def check_content_structure(self):
        return self.query("Состав содержания")

    def check_navigation(self):
        return self.query("Навигация")

    def check_headings(self):
        return self.query("Заголовки")

    def check_numbering(self):
        return self.query("Нумерация")

    # Структурные элементы
    def check_sections_and_subsections(self):
        return self.query("Разделы и подразделы")

    def check_abbreviations_list(self):
        return self.query("Список сокращений")

    def check_terms_list(self):
        return self.query("Список терминов")

    def check_introduction_and_conclusion(self):
        return self.query("Введение и заключение")

    def check_sources_list(self):
        return self.query("Список источников")

    def check_appendices(self):
        return self.query("Приложение")

    # Визуальные элементы
    def check_illustrations(self):
        return self.query("Иллюстрации")

    def check_tables(self):
        return self.query("Таблицы")

    # Стили
    def check_page_format(self):
        return self.query("Формат листа")

    def check_margins(self):
        return self.query("Размеры полей")

    def check_line_spacing(self):
        return self.query("Интервалы текста")

    def check_font_style(self):
        return self.query("Шрифт")
