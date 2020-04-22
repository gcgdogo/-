#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals

if __name__ =='__main__':
    import my_config

import kivy
from kivy.uix.screenmanager import ScreenManager, Screen , TransitionBase
from kivy.uix.button import Button

class Button_ScreenManager(ScreenManager):
    class Button_Screen(Screen):
        def __init__(self,Name_Self, ParrentScreenManager ):
            Screen.__init__(self , name=Name_Self)
            self.ParrentScreenManager = ParrentScreenManager

            self.Button = Button(text='')
            self.Button.bind(on_press=self.to_Button_Next)
            self.add_widget(self.Button)

        def to_Button_Next(self, *args):
            Button_text = args[0].text
            if self.ParrentScreenManager.Data_Source.Check_Text(Button_text):
                self.ParrentScreenManager.Switch_Button()
                self.ParrentScreenManager.on_Correct() #先触发Switch_Button 再 on_Correct


    def __init__(self , Data_Source , on_Correct = (lambda : None)):
        ScreenManager.__init__(self)
        TransitionBase.duration = 0.2  #设定翻页速度
        self.Data_Source = Data_Source
        self.on_Correct = on_Correct

        self.Button_Screen_1 = self.Button_Screen('Button_Screen_1' ,self)
        self.Button_Screen_2 = self.Button_Screen('Button_Screen_2' ,self)
        self.add_widget(self.Button_Screen_1)
        self.add_widget(self.Button_Screen_2)


    def Switch_Button(self , New_Text = None , Switch = True):
        if self.current == 'Button_Screen_1':
            Current_New = 'Button_Screen_2'
            Button_Screen_Old = self.Button_Screen_1
            Button_Screen_New = self.Button_Screen_2

        if self.current == 'Button_Screen_2':
            Current_New = 'Button_Screen_1'
            Button_Screen_Old = self.Button_Screen_2
            Button_Screen_New = self.Button_Screen_1

        if New_Text == None :    #允许通过直接指定的方式跳过获取文本
            New_Text = self.Data_Source.Next_Text()

        if Switch ==True:      #
            Button_Screen_New.Button.text = New_Text
            self.current = Current_New
        else:
            Button_Screen_Old.Button.text = New_Text

if __name__ =='__main__':
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout

    from Data_Source import Data_Source

    class MyApp(App):
        def build(self):
            self.Data_Source = Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4)

            List_BSM = []
            for I in range(4):
                List_BSM.append(Button_ScreenManager(self.Data_Source))

            layout = BoxLayout(orientation='vertical')
            for I_BSM in List_BSM:
                layout.add_widget(I_BSM)
            
            for I in zip(List_BSM , self.Data_Source.PreLoad_Text()):
                I[0].Switch_Button(New_Text = I[1] , Switch = False)

            return layout
    MyApp().run()