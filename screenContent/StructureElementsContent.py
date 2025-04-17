from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.anchorlayout import AnchorLayout

from components.ResultTable import ResultTable

class StructureElementsContent(AnchorLayout):

    def __init__(self, checker, **kwargs):
        super().__init__(**kwargs)
        self.anchor_x = 'center'
        self.anchor_y = 'top'

        self.checker = checker

        layout = BoxLayout(orientation='vertical', spacing=20,
                           size_hint=(None, None), pos_hint={'center_x': 0.5})
        layout.width = 600
        layout.height = 380

        layout.add_widget(Label(
            text="Структурные элементы",
            font_size=32,
            bold=True,
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=50
        ))

        layout.add_widget(Label(
            text="Результаты проверки разделов и подразделов шаблона ВКР.",
            font_size=18,
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=30
        ))

        checks = {
            "Разделы и подразделы": checker.check_sections_and_subsections,
            "Список сокращений": checker.check_abbreviations_list,
            "Список терминов": checker.check_terms_list,
            "Введение и заключение": checker.check_introduction_and_conclusion,
            "Список источников": checker.check_sources_list,
            "Приложение": checker.check_appendices,
        }

        self.result_table = ResultTable(checks)
        layout.add_widget(self.result_table)

        self.add_widget(layout)

    def update_checks(self):
        self.checks = {
            "Разделы и подразделы": self.checker.check_sections_and_subsections,
            "Список сокращений": self.checker.check_abbreviations_list,
            "Список терминов": self.checker.check_terms_list,
            "Введение и заключение": self.checker.check_introduction_and_conclusion,
            "Список источников": self.checker.check_sources_list,
            "Приложение": self.checker.check_appendices,
        }
        self.result_table.update_results(self.checks)