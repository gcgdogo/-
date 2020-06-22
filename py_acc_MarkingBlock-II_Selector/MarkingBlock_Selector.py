

################################################
#  函数定义
################################################

def func_YM_Cal( YM_input , X_offset ):
    Months = (YM_input // 100) * 12 + (YM_input % 100)
    Months = Months + X_offset - 1
    return (Months // 12) * 100 + (Months % 12) + 1


def func_Search_Value(sel_X , YM):
    for i in sel_X:
        if YM>=i[0] and YM<=i[1] :
            return i[2]
    return None


def func_FormatFactor(factor_x):
    #True 不变
    if factor_x == True:
        return True
    #字符串放到集合里
    if isinstance(factor_x , str):
        return set([factor_x])
    #能迭代直接转集合
    if '__iter__' in dir(factor_x):
        return set(factor_x)
    #剩下的直接转str吧
    return set([str(factor_x)])


def func_Selection_Compress(sel_X):
    sel_Temp = []
    for i in sel_X:
        #判定是否需要延长
        if len(sel_Temp) > 0:
            flag_extend = ( func_YM_Cal(sel_Temp[-1][1],1) == i[0] and sel_Temp[-1][2] == i[2])
        else:
            flag_extend = False

        #进行操作
        if flag_extend:
            sel_Temp[-1][1] = i[1]
        else:
            sel_Temp.append(i)
    
    return sel_Temp



def func_union(func_type , sel_A , sel_B , ignore_value = False):
    """func_type = ['交集','并集','减掉']"""
    sel_A.sort()
    sel_B.sort()
    
    #先批量获取起点和终点
    points_start = []
    points_end = []

    for i in sel_A + sel_B :
        points_start.append(i[0])
        points_end.append(func_YM_Cal(i[0],-1))

        points_start.append(func_YM_Cal(i[1],1))
        points_end.append(i[1])
    
    points_start = list(set(points_start))
    points_start.sort()
    points_start = points_start[:-1]  #去掉最后一个st

    points_end = list(set(points_end))
    points_end.sort()
    points_end = points_end[1:]  #去掉第一个ed
    
    sel_Temp = []
    for i_st , i_ed in zip(points_start , points_end):
        values = [
            func_Search_Value(sel_A , i_st) ,
            func_Search_Value(sel_B , i_st) ,
        ]
        #根据类型进行比较
        if func_type == '并集':
            check_data = (values[0] != None) or (values[1] != None)
        if func_type == '交集':
            check_data = ((values[0] != None) and (values[1] != None))
        if func_type == '减掉':
            check_data = ((values[0] != None) and (values[1] == None))

        if check_data :  #符合条件就生成值
            values_got = []
            for i in values:
                if i != None and i != True and not(i in values_got):
                    values_got.append(i)
            if len(values_got) == 0:
                if True in values :
                    values_got.append(True)
            
            #如果有值就开始添加
            if len(values_got) > 0 :
                if ignore_value :
                    val_got = True
                else:
                    if len(values_got) == 1 :
                        val_got = values_got[0]
                    else:
                        val_got = '#Fail{}'.format(values_got)
                sel_Temp.append([i_st , i_ed , val_got])
    
    sel_Temp = func_Selection_Compress(sel_Temp)
    print('A:{} \nB:{} \n=:{}'.format(sel_A,sel_B,sel_Temp))
    return sel_Temp




################################################
#  MarkingBlock_Selector
################################################

class MarkingBlock_Selector():
    """
    True 作为通配符，只要有这个项目就匹配（相当于原来的#）
    None 作为没有值的运算标记

    错误值统一用#Fail开头
    """
    def __init__(self , dict_data , selection = True , expression = '(ALL)' , math_lv = 'st'):
        """
        math_lv = ['st','+-','*','call']
        """
        self.data = dict_data
        #print(self.data.values())
        self.selection = selection  #[[st,ed,val=True]]
        if self.selection == True :
            #最大化选择
            self.selection = []
            for i in self.data.values():
                self.selection = func_union('并集',self.selection , i , ignore_value=True)
        
        self.expression = expression
        self.math_lv = math_lv

    #定义一个str用于打印
    def __str__(self , *args):
        #print('Call::__str__')
        return '<MarkingBlock_Selector: {} = {}>'.format(self.expression , self.selection)

    __repr__ = __str__
    __format__ = __str__

    #实现小括号
    def __call__(self , *args , **kargs):  
        #args不写等号，直接作为True判定
        for i in args :
            if not i in kargs :
                kargs[i] = True
        
        new_selection = self.selection.copy()

        if self.math_lv == 'st' :
            new_expression = '('
        if self.math_lv in ['+-','*']:
            new_expression = '( ' + self.expression + ' )('
        if self.math_lv == 'call':
            new_expression = self.expression + '('

        for i in kargs:
            kargs[i] = func_FormatFactor(kargs[i])
            new_expression = new_expression + '{}={},'.format(i,kargs[i])
        new_expression = new_expression[:-1] + ')'

        for i in kargs:
            sel_Temp = []
            if i in self.data:
                if kargs[i] == True :
                    sel_Temp = self.data[i]
                else :
                    for j in self.data[i]:
                        if j[2] in kargs[i] :
                            sel_Temp.append(j)
            new_selection = func_union('交集',new_selection , sel_Temp , ignore_value=True)

        #执行完成后返回全新的对象
        return MarkingBlock_Selector(
            dict_data= self.data ,
            selection= new_selection,
            expression=new_expression,
            math_lv='call'
        )

    #实现中括号取值
    def __getitem__(self , key):

        if self.math_lv in ['+-','*']:
            new_expression = '( ' + self.expression + ' )[{}]'.format(key)
        if self.math_lv == 'call':
            new_expression = self.expression + '[{}]'.format(key)

        #先检测一下有没有
        if not key in self.data:
            return MarkingBlock_Selector(
            dict_data= self.data ,
            selection= [],
            expression=new_expression,
            math_lv='call'
        )
        new_selection = func_union('交集' , self.selection , self.data[key])  #保留数值取交集

        #执行完成后返回全新的对象
        return MarkingBlock_Selector(
            dict_data= self.data ,
            selection= new_selection,
            expression=new_expression,
            math_lv='call'
        )

    #实现中括号赋值
    def __setitem__(self , key , source_data):
        #先判断一下有没有得到source_data
        if 'selection' in dir(source_data):
            if key in self.data:
                new_selection = self.data[key]
            else:
                new_selection = []
            new_selection = func_union('减掉',new_selection , source_data.selection)
            new_selection = func_union('并集',new_selection , source_data.selection)
            self.data[key] = new_selection
        #else:
            #new_selection = 

if __name__ == '__main__':
    x = MarkingBlock_Selector(
        {
            'a':[[200001,200112,'a']] ,
            'b':[[200005,200006,'b'],[200101,200105,'a'],[200111,202001,'c']]
        }
    )
    print('--------初始化完毕-------')
    print(x('b' ,a='a')['b'])
