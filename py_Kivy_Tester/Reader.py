#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals

if __name__ =='__main__':
    import my_config


from kivy.uix.rst import RstDocument
import time

class Reader(RstDocument):
    def __init__(self , Doc_Source , Scroll_TargetHeight = "Bottom" , on_DoubleClick = (lambda : None)):
        RstDocument.__init__(self , text=" " , base_font_size = 20)
        self.Doc_Source = Doc_Source
        self.Scroll_TargetHeight = Scroll_TargetHeight

        self.on_DoubleClick = on_DoubleClick
        self.List_ClickTime = []

        self.bind(on_touch_up = self.Check_DoubleClick , on_touch_down = self.Check_DoubleClick)

    def Get_Height(self):
        return self.viewport_size[1]

    def Update(self):
        self.text = self.Doc_Source()

        if self.Scroll_TargetHeight == "Bottom" :
            self.scroll_y = 0
        else:
            self.scroll_y = 1 - ( self.Scroll_TargetHeight() - self.size[1] / 2 ) / ( self.Get_Height() - self.size[1] )

    def Check_DoubleClick(self , *args):
        Val_Now = time.time()
        self.List_ClickTime.append(Val_Now)

        for I in range( len(self.List_ClickTime)-1 , -1 , -1):
            if Val_Now - self.List_ClickTime[I] > 0.5:
                self.List_ClickTime.pop(I)

        #print(list(map(lambda x : Val_Now - x , self.List_ClickTime)))

        if len(self.List_ClickTime) >= 4 :  #按下抬起时间都计算在内 总计达4个就行
            self.on_DoubleClick()




if __name__ =='__main__':
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout

    from Data_Source import Data_Source

    class MyApp(App):
        def build(self):
            self.Data_Source = Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4)

            self.Reader = Reader(Doc_Source=(lambda : "123456789"))

            self.Reader.Update()

            return self.Reader
    MyApp().run()