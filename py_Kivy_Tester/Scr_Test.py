#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals

if __name__ =='__main__':
    import my_config

from kivy.uix.screenmanager import Screen
from kivy.uix.boxlayout import BoxLayout
from Reader import Reader
from Button_ScreenManager import Button_ScreenManager

class Scr_Test(Screen):
    def __init__(self, ParrentScreenManager, Data_Source ):
        Screen.__init__(self , name = 'Scr_Test')
        self.ParrentScreenManager = ParrentScreenManager
        self.Data_Source = Data_Source


        #Reader         ########################
        self.Reader = Reader(
            Doc_Source = self.Data_Source.Visable_Doc,
            on_DoubleClick = self.ParrentScreenManager.Swich_Screen
        )
        self.Reader.size_hint = (0.7,1)
        self.Reader.Update()


        # Button_Layout ########################
        self.List_BSM = []
        for I in range(4):
            self.List_BSM.append(Button_ScreenManager(Data_Source = self.Data_Source , on_Correct = self.Reader.Update ))

        self.Button_Layout = BoxLayout(orientation='vertical')
        self.Button_Layout.size_hint = (0.3,1)

        for I_BSM in self.List_BSM:
            self.Button_Layout.add_widget(I_BSM)
        
        for I in zip(self.List_BSM , self.Data_Source.PreLoad_Text()):
            I[0].Switch_Button(New_Text = I[1] , Switch = False)


        #Layout         ########################
        self.Layout = BoxLayout(orientation='horizontal')
        self.Layout.add_widget(self.Reader)
        self.Layout.add_widget(self.Button_Layout)

        self.add_widget(self.Layout)

        #设定切换事件
        self.bind(on_pre_enter = self.on_pre_enter , on_enter = self.on_enter)

    def on_pre_enter(self , *args):  #清空选项
        for I in self.List_BSM:
            I.Switch_Button(New_Text = " ", Switch = False)

    def on_enter(self , *args):  #载入选项
        for I in zip(self.List_BSM , self.Data_Source.PreLoad_Text()):
            I[0].Switch_Button(New_Text = I[1] , Switch = False)

if __name__ =='__main__':
    from kivy.app import App

    from Data_Source import Data_Source
    from kivy.uix.screenmanager import ScreenManager

    class MyApp(App):
        def build(self):
            self.Data_Source = Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4)
            self.ScreenManager = ScreenManager()
            self.ScreenManager.add_widget(Scr_Test(self.Data_Source))
            return self.ScreenManager
    MyApp().run()