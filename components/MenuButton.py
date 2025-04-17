from kivy.uix.button import Button
from kivy.core.window import Window

class MenuButton(Button):
    def __init__(self, screen_manager=None, target_screen=None, **kwargs):
        super().__init__(**kwargs)
        self.screen_manager = screen_manager
        self.target_screen = target_screen

        self.background_normal = ''
        self.background_color = (0, 0, 0, 0)
        self.color = (0, 0, 0, 1)
        Window.bind(mouse_pos=self.on_mouse_pos)
        self.hover = False

        self.bind(on_release=self.switch_screen)

    def on_mouse_pos(self, window, pos):
        if not self.get_root_window():
            return

        inside = self.collide_point(*self.to_widget(*pos))
        if inside and not self.hover:
            self.hover = True
            self.background_color = (1, 1, 1, 1)
        elif not inside and self.hover:
            self.hover = False
            self.background_color = (0, 0, 0, 0)

    def switch_screen(self, *args):
        if self.screen_manager and self.target_screen:
            self.screen_manager.current = self.target_screen
