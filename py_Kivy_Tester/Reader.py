#coding:UTF-8

if __name__ =='__main__':
    import my_config


from kivy.uix.rst import RstDocument

class Reader(RstDocument):
    def __init__(self , Doc_Source , Scroll_TargetHeight = "Bottom"):
        RstDocument.__init__(self , text=" " , base_font_size = 20)
        self.Doc_Source = Doc_Source
        self.Scroll_TargetHeight = Scroll_TargetHeight

    def Get_Height(self):
        return self.viewport_size[1]

    def Update(self):
        self.text = self.Doc_Source()

        if self.Scroll_TargetHeight == "Bottom" :
            self.scroll_y = 0
        else:
            self.scroll_y = 1 - ( self.Scroll_TargetHeight() - self.size[1] / 2 ) / ( self.Get_Height() - self.size[1] )

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