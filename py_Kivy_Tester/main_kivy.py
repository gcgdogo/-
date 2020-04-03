#coding:UTF-8
from __future__ import division,print_function,absolute_import


import my_config
from kivy.app import App

from Base_ScrMan import Base_ScrMan


class MyApp(App):
    def build(self):
        return Base_ScrMan()
        
MyApp().run()
