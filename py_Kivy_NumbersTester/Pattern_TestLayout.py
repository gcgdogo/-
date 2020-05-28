#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals

if __name__ =='__main__':
    import my_config

import kivy
from kivy.clock import Clock

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from random import randint

import Pattern_ScreenManager
import Pattern_Libary

import time

class Pattern_TestLayout(BoxLayout):
    def __init__(self):
        BoxLayout.__init__(self,orientation='horizontal')
        self.UpdatedSize = (0,0)

        self.QuestLayout = BoxLayout(orientation='horizontal' , size_hint_x = 0.8)
        self.QuestPattern = Pattern_ScreenManager.Pattern_ScreenManager()
        self.QuestPattern .size_hint_x = 0.5
        
        self.QuestLayout.add_widget(self.QuestPattern)
        self.QuestPattern.Switch_Button(Pattern_Libary.Rand_Pattern(randint(1,3)))

        self.ChoiceLayout = BoxLayout(orientation='vertical' , size_hint_x = 0.2)
        self.List_ChoicePattern = []
        for I in range(3) :
            self.List_ChoicePattern.append(Pattern_ScreenManager.Pattern_ScreenManager())

            self.ChoiceLayout.add_widget(self.List_ChoicePattern[I])
            self.List_ChoicePattern[I].Switch_Button(Pattern_Libary.Rand_Pattern(I+1))
            self.List_ChoicePattern[I].press_target = self.Choice_Pressed
        
        self.add_widget(self.QuestLayout)
        self.add_widget(self.ChoiceLayout)

        self.List_ToSwitch = []

        Clock.schedule_interval(self.Check_Size, 1)



    def Check_Size(self,*args):
        if self.UpdatedSize != (self.width,self.height) :
            self.ChoiceLayout.width = self.height / 3
            self.ChoiceLayout.size_hint_max = [self.height / 3 , self.height]

            Quest_Size = min(self.width - self.height / 3 , self.height)
            self.QuestPattern.size = [Quest_Size , Quest_Size]
            self.QuestPattern.size_hint_max = [Quest_Size , Quest_Size]
            self.QuestPattern.pos[0] = max((self.width - self.height / 3 - Quest_Size)/2,0)

            self.UpdatedSize = (self.width,self.height)



    def Choice_Pressed(self):
        Choose_Right = False
        for I in range(len(self.List_ChoicePattern)):
            if self.List_ChoicePattern[I].is_pressed :
                if self.List_ChoicePattern[I].Get_Value() ==  self.QuestPattern.Get_Value():
                    self.List_ChoicePattern[I].Set_PatternBackground('Green')
                    Choose_Right = True
                else:
                    self.List_ChoicePattern[I].Set_PatternBackground('Red')
        
        if Choose_Right :
            self.List_ToSwitch.append(
                (self.QuestPattern , Pattern_Libary.Rand_Pattern(randint(1,3)))
            )
            for I in range(len(self.List_ChoicePattern)):
                if self.List_ChoicePattern[I].is_pressed :
                    self.List_ToSwitch.append(
                        (self.List_ChoicePattern[I] , Pattern_Libary.Rand_Pattern(I+1))
                    )
                #self.List_ChoicePattern[I].Switch_If_Pressed(Pattern_Libary.Rand_Pattern(I+1))  #旧版 
            self.Switch_ByTurn()

    def Switch_ByTurn(self , *args):
        #[(SM1,PattData1),,,]
        if len(self.List_ToSwitch) > 0 :
            (Tar_SM , Tar_PattData) = self.List_ToSwitch[0]
            Tar_SM.Switch_Button(Tar_PattData)
            self.List_ToSwitch = self.List_ToSwitch[1:]
            Clock.schedule_once(self.Switch_ByTurn , 0.21)

if __name__ =='__main__':
    from kivy.app import App

    class MyApp(App):
        def build(self):
            Layout = Pattern_TestLayout()
            return Layout
    MyApp().run()