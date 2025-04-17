from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.popup import Popup
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.anchorlayout import AnchorLayout

class MainScreenContent(AnchorLayout):

    def __init__(self, screen_manager, checker, **kwargs):
        super().__init__(**kwargs)
        self.anchor_x = 'center'
        self.anchor_y = 'top'
        self.screen_manager = screen_manager
        self.checker = checker


        layout = BoxLayout(orientation='vertical', spacing=20, size_hint=(None, None), pos_hint={'center_x': 0.5})
        layout.width = 600
        layout.height = 300

        layout.add_widget(Label(
            text="Главная",
            font_size=32,
            bold=True,
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=50
        ))

        layout.add_widget(Label(
            text="Пожалуйста, прикрепите файл шаблона ВКР для проверки.",
            font_size=18,
            color=(0, 0, 0, 1),
            size_hint=(1, None),
            height=30
        ))

        button_layout = BoxLayout(orientation='horizontal', spacing=20, size_hint=(None, None), pos_hint={'center_x': 0.5})
        button_layout.width = 500

        attach_button = Button(
            text="Прикрепить файл",
            size_hint=(None, None),
            size=(250, 70),
            background_color=(0.886, 0.886, 0.886, 1),
            color=(1, 1, 1, 1)
        )
        attach_button.bind(on_press=self.open_file_chooser)

        submit_button = Button(
            text="Отправить на проверку",
            size_hint=(None, None),
            size=(250, 70),
            background_color=(0.886, 0.886, 0.886, 1),
            color=(1, 1, 1, 1)
        )
        submit_button.bind(on_press=self.send_for_check)

        button_layout.add_widget(attach_button)
        button_layout.add_widget(submit_button)

        self.status_label = Label(
            text="",
            font_size=18,
            size_hint=(1, None),
            height=30,
            color=(0, 0, 0, 0)
        )

        layout.add_widget(button_layout)
        layout.add_widget(self.status_label)

        self.add_widget(layout)

    def open_file_chooser(self, instance):
        global selected_file
        filechooser = FileChooserIconView()
        popup = Popup(
            title="Выберите файл",
            content=filechooser,
            size_hint=(0.9, 0.9)
        )

        def select_file(filechooser, selection, touch):
            if selection:
                App.get_running_app().selected_file = selection[0]
                print("Файл выбран:", App.get_running_app().selected_file)
                popup.dismiss()

        filechooser.bind(on_submit=select_file)
        popup.open()

    def send_for_check(self, instance):
        if App.get_running_app().selected_file:
            print(f"Файл {App.get_running_app().selected_file} отправлен на проверку")
            for screen in self.screen_manager.screens:
                if hasattr(screen, 'update_content_checks') and callable(screen.update_content_checks):
                    screen.update_content_checks()
        else:
            print("Файл не выбран")

        result = self.checker.run_all_checks()

        if result:
            self.status_label.text = "Проверка пройдена успешно"
            self.status_label.color = (0, 0.6, 0, 1)
        else:
            self.status_label.text = "Шаблон не прошел проверку"
            self.status_label.color = (1, 0, 0, 1)