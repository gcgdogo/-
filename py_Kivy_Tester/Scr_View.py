#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals
if __name__ =='__main__':
    import my_config

from kivy.uix.screenmanager import Screen
from kivy.uix.boxlayout import BoxLayout
from Reader import Reader
from kivy.uix.slider import Slider

class View_Slider(Slider):
    def __init__(self, Data_Source , on_Value_Change = (lambda : None)):
        Slider.__init__(self)
        self.Data_Source = Data_Source
        self.on_Value_Change = on_Value_Change
        self.bind(on_touch_move = self.Value_Change , on_touch_up = self.Value_Change)

        self.size_hint_max_y = 60


    def Value_Refresh(self , *args):
        self.min = 0
        self.max = self.Data_Source.Finger_Max
        self.value = self.Data_Source.Finger

    def Value_Change(self , *args):
        if self.Data_Source.Finger == int(self.value):  #没有移动就不改啦
            return
        self.Data_Source.Finger = int(self.value)
        self.on_Value_Change()
        #print(self.Data_Source.Whole_Doc())


class Scr_View(Screen):
    def __init__(self, ParrentScreenManager , Data_Source , Doc_Source , Scroll_TargetHeight = "Bottom"):
        if Scroll_TargetHeight == "Bottom" :
            Screen.__init__(self , name = 'Scr_View_Hidden')
        else:
            Screen.__init__(self , name = 'Scr_View')
            
        self.ParrentScreenManager = ParrentScreenManager
        self.Data_Source = Data_Source
        self.Doc_Source = Doc_Source
    
        #Reader         ########################
        self.Reader = Reader(
            Doc_Source=self.Doc_Source , 
            Scroll_TargetHeight= Scroll_TargetHeight ,
            on_DoubleClick = self.ParrentScreenManager.Swich_Screen
        )
        self.Reader.Update()

        #Slider         ########################
        self.View_Slider = View_Slider(self.Data_Source , self.Reader.Update)

        self.Layout = BoxLayout(orientation='vertical')
        self.Layout.add_widget(self.Reader)
        self.Layout.add_widget(self.View_Slider)

        self.add_widget(self.Layout)
        #绑定事件
        self.bind(on_pre_enter = self.on_pre_enter)

    def on_pre_enter(self , *args):  #初始化
        self.View_Slider.Value_Refresh()
        self.Reader.Update()

if __name__ =='__main__':
    from kivy.app import App

    from Data_Source import Data_Source
    from kivy.uix.screenmanager import ScreenManager

    class MyApp(App):
        def build(self):
            self.Data_Source = Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4)
            self.ScreenManager = ScreenManager()
            self.ScreenManager.Swich_Screen = lambda : print('clicked')
            self.ScreenManager.add_widget(Scr_View(
                self.ScreenManager ,
                Data_Source = self.Data_Source ,
                Doc_Source = self.Data_Source.Whole_Doc
                ))
            return self.ScreenManager
    MyApp().run()