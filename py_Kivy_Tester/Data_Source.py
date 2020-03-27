#coding:UTF-8

import random

class Data_Source():
    def __init__(self, Data_FileName , N_Choices):
        self.Data_FileName = Data_FileName
        self.N_Choices = N_Choices
        self.Load_Data() #读取数据

        self.Finger = 0
        self.Finger_Max = len(self.List_Choices) - 1
        self.Arrage_Choices()  # 排序选项
        
    #重新读取数据 , 只读取 Data_FileName 不修改 N_Choices
    def Reload(self , Data_FileName):
        self.__init__(Data_FileName , self.N_Choices)

    #读取，排序
    def Load_Data(self):
        Data_Loaded = eval(
            open(self.Data_FileName,'r',encoding='UTF-8').read()
        )
        self.Source_Text = Data_Loaded['Source_Text']
        self.List_Choices = Data_Loaded['List_Choices']

    def Arrage_Choices(self):
        Stack_Fingers = []
        Stack_Open = False  #只要堆栈足够长，就开启堆栈出口
        for I in range(self.Finger_Max , - self.N_Choices , -1):
            if I >= 0 :
                Stack_Fingers.append(I)
            
            if len(Stack_Fingers) >= self.N_Choices :
                Stack_Open = True
            
            if Stack_Open:
                R_Int = random.randint(0,len(Stack_Fingers) - 1)
                #print('[I = {} , R_Int = {}] {}'.format(I , R_Int , Stack_Fingers))  #用于调试

                self.List_Choices[
                    Stack_Fingers.pop(R_Int)
                ][2] = I
        

    #获取初始选项
    def PreLoad_Text(self):
        List_PreLoad_Text = []
        for I in self.List_Choices[self.Finger : ]:
            if I[2] <= self.Finger :
                List_PreLoad_Text.append(I[0])
        return List_PreLoad_Text
    
    #选项文本检测及输出
    def Check_Text(self , x_Text):
        if self.Finger >= self.Finger_Max : return False  # 超出Finger_Max
        return x_Text == self.List_Choices[self.Finger][0]
    
    def Next_Text(self):
        self.Finger = self.Finger + 1
        for I in self.List_Choices[self.Finger : ]:
            if I[2] == self.Finger :
                return I[0]
        return '< End >'

    #文章输出相关准备
    def Doc_Colorize(self , str_Doc , str_Color):
        str_Doc_Colored = str_Doc
        str_Doc_Colored = str_Doc_Colored.replace(
            '\n' ,
            '[/color]\n[color={}]'.format(str_Color)
        )
        str_Doc_Colored = '[color={}]'.format(str_Color) + str_Doc_Colored + '[/color]'
        str_Doc_Colored = str_Doc_Colored.replace(
            '[color={}][/color]'.format(str_Color) ,
            ''
        )
        return str_Doc_Colored

    def Visable_Len(self):
        if self.Finger >= self.Finger_Max : return len(self.Source_Text)  # 超出Finger_Max
        return self.List_Choices[self.Finger][1]


    #文章输出
    def Visable_Doc(self):
        return self.Source_Text[:self.Visable_Len()]

    def Whole_Doc(self):
        return self.Visable_Doc() + self.Doc_Colorize(self.Source_Text[self.Visable_Len() : ] , '3333cc')

if __name__ =='__main__':
    print(Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4).Whole_Doc())