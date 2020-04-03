#coding:UTF-8
from __future__ import division,print_function,absolute_import

import random

class Data_Source():
    def __init__(self, Data_FileName , N_Choices):
        self.Data_FileName = Data_FileName
        self.N_Choices = N_Choices
        self.Load_Data() #读取数据

        self.Finger = 0
        self.Finger_Max = len(self.List_Choices) - 1
        self.Arrage_Choices()  # 排序选项

        self.DocColorizeStock_str_Doc = ""
        self.DocColorizeStock_str_Color = ""
        self.str_Doc_Colored = ""  #保存涂色文档结果
        
    #重新读取数据 , 只读取 Data_FileName 不修改 N_Choices
    def Reload(self , Data_FileName):
        self.__init__(Data_FileName , self.N_Choices)

    #读取，排序
    def Load_Data(self):
        Str_Loaded = open(self.Data_FileName,'rb').read().decode('UTF-8')
        #Str_Loaded_B = open(self.Data_FileName,'r',encoding='UTF-8').read()
        #Data_Loaded = eval(Str_Loaded_B)   #测试一下这种的能eval不
        Data_Loaded = eval(Str_Loaded)

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
        if self.DocColorizeStock_str_Doc == str_Doc and self.DocColorizeStock_str_Color == str_Color:
            return self.str_Doc_Colored   #没有调整就不重新上色了

        Len_InLine = str_Doc.find('\n')
        #print('Len_InLine = {}'.format(Len_InLine))

        str_InLine = str_Doc[:Len_InLine]
        str_NextLine = str_Doc[Len_InLine:]
        
        str_InLine = '[/color][color={}]'.format(str_Color).join(
            list(str_InLine)
        )
        str_NextLine = str_NextLine.replace(
            '\n' ,
            '[/color]\n[color={}]'.format(str_Color)
        )

        #print('str_InLine = {}'.format(str_InLine))
        #print('str_NextLine = {}'.format(str_NextLine))

        self.str_Doc_Colored = '[color={}]'.format(str_Color) + str_InLine + str_NextLine + '[/color]'

        #print('self.str_Doc_Colored = {}'.format(self.str_Doc_Colored))

        self.str_Doc_Colored = self.str_Doc_Colored.replace(
            '[color={}][/color]'.format(str_Color) ,
            ''
        )

        #print('self.str_Doc_Colored = {}'.format(self.str_Doc_Colored))
        return self.str_Doc_Colored

    def Visable_Len(self):
        if self.Finger >= self.Finger_Max : return len(self.Source_Text)  # 超出Finger_Max
        return self.List_Choices[self.Finger][1]


    #文章输出
    def Visable_Doc(self):
        str_Visable_Doc = self.Source_Text[:self.Visable_Len()]
        print('[Visable_Doc] :: ' + str_Visable_Doc)
        return str_Visable_Doc

    def Whole_Doc(self):
        str_Whole_Doc = self.Visable_Doc() + self.Doc_Colorize(self.Source_Text[self.Visable_Len() : ] , '3333cc')
        print('[Whole_Doc] :: ' + str_Whole_Doc)
        return str_Whole_Doc

if __name__ =='__main__':
    print(Data_Source('中国强化知识产权保护为创新发展“护航”.repr.txt',4).Whole_Doc())