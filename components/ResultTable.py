from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout

class ResultTable(GridLayout):
    def __init__(self, checks: dict[str, callable], **kwargs):
        super().__init__(cols=2, spacing=10, size_hint=(1, None), **kwargs)
        self.bind(minimum_height=self.setter('height'))
        self.populate(checks)

    def populate(self, checks: dict[str, callable]):
        self.clear_widgets()

        for criterion, check_function in checks.items():
            try:
                passed = check_function()
                if passed is None:
                    result_text = "Нет результатов"
                    result_color = (0.5, 0.5, 0.5, 1)
                else:
                    result_text = "Успешно" if passed else "Провалено"
                    result_color = (0, 0.6, 0, 1) if passed else (1, 0, 0, 1)
            except Exception as e:
                result_text = "Ошибка"
                result_color = (1, 0, 0, 1)
                print(f"Ошибка в проверке '{criterion}':", e)

            self.add_widget(Label(
                text=criterion,
                font_size=18,
                color=(0, 0, 0, 1),
                size_hint_y=None,
                height=30,
                halign='left',
                valign='middle'
            ))

            self.add_widget(Label(
                text=result_text,
                font_size=18,
                color=result_color,
                size_hint_y=None,
                height=30,
                halign='right',
                valign='middle'
            ))

    def update_results(self, checks: dict[str, callable]):
        self.populate(checks)