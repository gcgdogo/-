#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals


import my_config
from kivy.app import App

from Pattern_TestLayout import Pattern_TestLayout


class MyApp(App):
    def build(self):
        return Pattern_TestLayout()
        
MyApp().run()
