from kivy.uix.screenmanager import Screen

from screenContent.MainScreenContent import MainScreenContent

class MainScreen(Screen):
    def __init__(self, screen_manager, checker, **kwargs):
        super().__init__(**kwargs)

        self.add_widget(MainScreenContent(screen_manager, checker))

    def update_bg(self, *args):
        self.bg.pos = self.pos
        self.bg.size = self.size
