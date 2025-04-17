from kivy.uix.screenmanager import Screen

from screenContent.StylesContent import StylesContent

class StylesScreen(Screen):
    def __init__(self, checker, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'horizontal'

        self.content_widget = StylesContent(checker=checker)
        self.add_widget(self.content_widget)

    def update_content_checks(self):
        self.content_widget.update_checks()

    def update_bg(self, *args):
        self.bg.pos = self.pos
        self.bg.size = self.size