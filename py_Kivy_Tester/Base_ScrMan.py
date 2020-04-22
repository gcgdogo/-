#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals
if __name__ =='__main__':
    import my_config
from kivy.app import App

from Data_Source import Data_Source
from kivy.uix.screenmanager import ScreenManager , SlideTransition

from Scr_Test import Scr_Test
from Scr_View import Scr_View

class Base_ScrMan(ScreenManager):
    def __init__(self):
        ScreenManager.__init__(self)
        self.Data_Source = Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4)
        self.Scr_Test = Scr_Test(
            self , 
            self.Data_Source
        )

        self.Scr_View_Hidden = Scr_View(
            self , 
            Data_Source = self.Data_Source , 
            Doc_Source = self.Data_Source.Visable_Doc
        )

        self.Scr_View = Scr_View(
            self , 
            Data_Source = self.Data_Source , 
            Doc_Source = self.Data_Source.Whole_Doc ,
            Scroll_TargetHeight = self.Scr_View_Hidden.Reader.Get_Height 
        )

        self.add_widget(self.Scr_Test)
        self.add_widget(self.Scr_View)
        self.add_widget(self.Scr_View_Hidden)

        self.current = 'Scr_Test' #初始化屏幕

    def Swich_Screen(self , Target_Screen = '' ):
        if Target_Screen == '':
            if self.current == 'Scr_Test' : Target_Screen = 'Scr_View'
            if self.current == 'Scr_View' : Target_Screen = 'Scr_Test'
        
        Slide_Direction = 'left'
        if self.current == 'Scr_View' : Slide_Direction = 'right'

        self.transition.direction = Slide_Direction
        self.current = Target_Screen

class MyApp(App):
    def build(self):
        return Base_ScrMan()
MyApp().run()