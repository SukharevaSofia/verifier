from kivy.uix.boxlayout import BoxLayout
from kivy.uix.image import Image
from kivy.uix.widget import Widget
from kivy.graphics import Color, Rectangle
from components.MenuButton import MenuButton

class LeftMenu(BoxLayout):
    def __init__(self, screen_manager, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.size_hint_x = None
        self.width = 300

        with self.canvas.before:
            Color(0.898, 0.898, 0.898, 1)
            self.bg = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self.update_bg, size=self.update_bg)

        self.add_widget(Image(source='pict/ITMO_logo.png', size_hint_y=None, height=100))
        self.add_widget(Widget(size_hint_y=None, height=10))

        menu_items = [
            ("Главная", "main"),
            ("Содержание ВКР", "content"),
            ("Структурные элементы", "structure"),
            ("Визуальные элементы", "visual"),
            ("Стили", "styles")
        ]

        for label, screen_name in menu_items:
            btn = MenuButton(text=label, screen_manager=screen_manager, target_screen=screen_name)
            btn.size_hint_y = None
            btn.height = 50
            self.add_widget(btn)
            self.add_widget(Widget(size_hint_y=None, height=10))

        self.add_widget(Widget())

    def update_bg(self, *args):
        self.bg.pos = self.pos
        self.bg.size = self.size