from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.anchorlayout import AnchorLayout

from components.ResultTable import ResultTable

class VisualElementsContent(AnchorLayout):

    def __init__(self, checker, **kwargs):
        super().__init__(**kwargs)
        self.anchor_x = 'center'
        self.anchor_y = 'top'
        self.checker = checker

        layout = BoxLayout(orientation='vertical', spacing=20,
                           size_hint=(None, None), pos_hint={'center_x': 0.5})
        layout.width = 600
        layout.height = 220

        layout.add_widget(Label(
            text="Визуальные элементы",
            font_size=32,
            bold=True,
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=50
        ))

        layout.add_widget(Label(
            text="Результаты проверки таблиц и иллюстраций.",
            font_size=18,
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=30
        ))

        checks = {
            "Иллюстрации": checker.check_illustrations,
            "Таблицы": checker.check_tables
        }

        self.result_table = ResultTable(checks)
        layout.add_widget(self.result_table)

        self.add_widget(layout)

    def update_checks(self):
        self.checks = {
            "Иллюстрации": self.checker.check_illustrations,
            "Таблицы": self.checker.check_tables,
        }
        self.result_table.update_results(self.checks)