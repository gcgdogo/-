#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals

if __name__ =='__main__':
    import my_config

import kivy
from kivy.clock import Clock
from kivy.uix.screenmanager import ScreenManager, Screen , TransitionBase
from kivy.uix.button import Button
from kivy.graphics import Color, Rectangle , Ellipse

from math import sin , cos , pi

class PatternButton(Button):
    def __init__(self):
        Button.__init__(self , text = "不应该能看到这个文字啊")

        self.PatternData = {
            'value' : 0 ,
            'nodes' :[
                {
                    'color' : [0,0,0],
                    'shape' : 'circle'
                }
            ]
        }

        self.UpdatedSize = (0,0)
        self.UpdatePattern()
        self.bind(on_press = self.UpdatePattern)

        Clock.schedule_interval(self.CheckSize, 1)

    def UpdatePattern(self , *args):
        
        #用于判定形式，取齐长度
        def Value_To_List(x , y):
            if not(isinstance(x,list)) : x = [x]
            if not(isinstance(y,list)) : y = [y]
            len_xy = max(len(x),len(y))

            x = (x*len_xy)[:len_xy]
            y = (y*len_xy)[:len_xy]

            return list(zip(x,y))


        self.Draw_Color = [0,0,0]
        self.Draw_Shape = 'circle'
        self.Draw_Size = 1

        self.canvas.clear()  #先清空再作图
        self.canvas.children = []

        self.canvas.add(Color(1,1,1))  #白色背景
        self.canvas.add(Rectangle(pos = [self.x + 2 , self.y + 2] , size = [self.width - 4,self.height - 4]))

        for Node_Data in self.PatternData['nodes']:
            #print(Node_Data)

            #如果有前置信息就直接读取
            if 'color' in Node_Data :
                self.Draw_Color = Node_Data['color']
            if 'shape' in Node_Data:
                self.Draw_Shape = Node_Data['shape']
            
            #判定XY模式
            if 'x' in Node_Data and 'y' in Node_Data :
                list_xy = Value_To_List(Node_Data['x'] , Node_Data['y'])
                for x , y in list_xy:
                    self.Draw('xy',x,y)

            #rt模式  极坐标
            if 'r' in Node_Data and 't' in Node_Data :
                list_rt = Value_To_List(Node_Data['r'] , Node_Data['t'])
                for r , t in list_rt:
                    self.Draw('rt',r,t)

        #print(len(self.canvas.children))
        self.UpdatedSize = (self.width,self.height)


    def Draw(self , Mode = 'xy' , x = 0 , y = 0):
        #print('x = {} , y = {}'.format(x,y))
        if Mode == 'rt':
            r = x
            t = y * 2 * pi

            x = r * sin(t)
            y = r * cos(t)

        Canvas_Size = min(self.width , self.height) * 0.8
        size = Canvas_Size/8    #大小先按1/8定

        x = x * Canvas_Size/2 + self.center_x - size/2
        y = y * Canvas_Size/2 + self.center_y - size/2

        #print(self.Draw_Color.rgb)

        self.canvas.add(Color(*self.Draw_Color))  #必须用心的Color，怀疑对象里面有停止使用的标记，如果用同一个color，就会被停用了

        if self.Draw_Shape == 'circle' :
            #print('self.width ={} , self.height = {} , x = {} , y = {}'.format(self.width , self.height , x,y))
            self.canvas.add(Ellipse(pos = [x , y] , size = [size,size] , group = 'Draw' ,color = Color(0.5,0.5,0.5)))

    def CheckSize(self , *args):
        if self.UpdatedSize != (self.width,self.height):
            print(repr((self.width,self.height)))
            self.UpdatePattern()

#####################################################################################

class Pattern_ScreenManager(ScreenManager):
    class Pattern_Screen(Screen):
        #点击时触发，Func_Press_Target

        def __init__(self,Name_Self, Func_Press_Target = (lambda x : print('Func_Press :: ' + repr(x))) ):
            Screen.__init__(self , name=Name_Self)
            self.Func_Press_Target = Func_Press_Target
            
            self.PatternButton = PatternButton()
            self.PatternButton.bind(on_press=self.Func_Press)
            self.add_widget(self.PatternButton)

        def Func_Press(self , *args):
            self.Func_Press_Target(*args)

        def Edit_Pattern(self, PatternData):
            self.PatternButton.PatternData = PatternData
            self.PatternButton.UpdatePattern()


    def __init__(self):
        ScreenManager.__init__(self)
        TransitionBase.duration = 0.2  #设定翻页速度
        self.is_pressed = False
        self.press_target = (lambda *args : print(args))

        self.Pattern_Screen_1 = self.Pattern_Screen('Pattern_Screen_1' ,self.Func_Press)
        self.Pattern_Screen_2 = self.Pattern_Screen('Pattern_Screen_2' ,self.Func_Press)
        self.add_widget(self.Pattern_Screen_1)
        self.add_widget(self.Pattern_Screen_2)

        #设置一个默认图样吧，要不啥也看不到啊
        self.default_PatternData = {
            'value' : 0 ,
            'nodes' :[
                {
                    'color' : [0,0,1],
                    'shape' : 'circle'
                },
                {
                    'y' : [0] , 
                    'x' : [0]
                },
                {
                    'r' : [1,0.5] , 
                    't' : [0,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9]
                },

            ]
        }


    def Func_Press(self , *args):
        self.is_pressed = True
        self.press_target()

    def Switch_If_Pressed(self , PatternData = {}):
        if self.is_pressed :
            self.Switch_Button(PatternData)
            self.is_pressed = False

    def Switch_Button(self , PatternData = {} , Switch = True):
        if PatternData == {} : PatternData = self.default_PatternData

        if self.current == 'Pattern_Screen_1':
            Current_New = 'Pattern_Screen_2'
            Pattern_Screen_Old = self.Pattern_Screen_1
            Pattern_Screen_New = self.Pattern_Screen_2

        if self.current == 'Pattern_Screen_2':
            Current_New = 'Pattern_Screen_1'
            Pattern_Screen_Old = self.Pattern_Screen_2
            Pattern_Screen_New = self.Pattern_Screen_1

        if Switch ==True:
            Pattern_Screen_New.PatternButton.size = self.size  #大概翻转的时候才同步size，不提前设定size的话，第一次翻转的大小不正常
            Pattern_Screen_New.Edit_Pattern(PatternData)
            self.current = Current_New
        else:
            Pattern_Screen_Old.Edit_Pattern(PatternData)




if __name__ =='__main__':
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout

    import Pattern_Libary

    class MyApp(App):
        def build(self):
            layout = BoxLayout(orientation='vertical')
            self.list_PatternSM = []

            for I in range(3) :
                self.list_PatternSM.append(Pattern_ScreenManager())

                layout.add_widget(self.list_PatternSM[I])
                self.list_PatternSM[I].Switch_Button(Pattern_Libary.Rand_Pattern(I+1))
                self.list_PatternSM[I].press_target = self.press
                
            return layout
        
        def press(self):
            for I in range(len(self.list_PatternSM)):
                self.list_PatternSM[I].Switch_If_Pressed(Pattern_Libary.Rand_Pattern(I+1))

    MyApp().run()