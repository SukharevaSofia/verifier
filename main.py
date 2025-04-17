from kivy.app import App
from kivy.graphics import Color, Rectangle
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.screenmanager import ScreenManager

from logic.Checker import Checker
from screens.MainScreen import MainScreen
from screens.ContentWKRScreen import ContentWKRScreen
from screens.StructureElementsScreen import StructureElementsScreen
from screens.VisualElementsScreen import VisualElementsScreen
from screens.StylesScreen import StylesScreen

from components.LeftMenu import LeftMenu

class DesignApp(App):
    selected_file = None
    def build(self):
        root = BoxLayout(orientation='horizontal')

        checker = Checker()
        sm = ScreenManager()
        sm.add_widget(MainScreen(screen_manager=sm, checker=checker, name='main'))
        sm.add_widget(ContentWKRScreen(checker=checker, name='content'))
        sm.add_widget(StructureElementsScreen(checker=checker, name='structure'))
        sm.add_widget(VisualElementsScreen(checker=checker,name='visual'))
        sm.add_widget(StylesScreen(checker=checker,name='styles'))

        with sm.canvas.before:
            Color(1, 1, 1, 1)
            self.rect = Rectangle(size=sm.size, pos=sm.pos)

        sm.bind(size=self.update_rect, pos=self.update_rect)
        menu = LeftMenu(screen_manager=sm)

        root.add_widget(menu)
        root.add_widget(sm)

        return root

    def update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size
if __name__ == '__main__':
    DesignApp().run()