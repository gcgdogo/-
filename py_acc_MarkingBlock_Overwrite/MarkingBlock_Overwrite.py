import win32com.client

def YM_Cal( YM_input , X_offset ):
    Months = (YM_input // 100) * 12 + (YM_input % 100)
    Months = Months + X_offset - 1
    return (Months // 12) * 100 + (Months % 12) + 1

class DataSet_Origin:

    def __init__(self,ado_con,table_name,Key_NameList):
        #需要写入的行先临时保存起来
        self.temp_LinesToWrite = []

        #保存ado链接
        self.ado_con = ado_con

        #保存表名
        self.table_name = table_name

        #标准字段写入读取顺序
        self.Key_NameList = Key_NameList

        #计算SQL查询字符串
        #先把Key_NameList转为字符串
        str_Keys = ' , '.join(Key_NameList)
        str_SQL = 'select {} from {} order by 员工字段ID , 起始年月'.format(str_Keys,self.table_name)

        #建立Recordset链接
        self.ado_rs = win32com.client.Dispatch(r'ADODB.Recordset')
        self.ado_rs.ActiveConnection = self.ado_con
        self.ado_rs.Source = str_SQL
        self.ado_rs.CursorType = 2
        self.ado_rs.LockType = 3
        self.ado_rs.Open()

        #读取第一行
        self.line = [] #空行用空列表[]表示
        self.line_LoadNext(LoadFirstLine = True)

        self.line_Edited = False

        #保存顺序位置
        self.index_dict = {}
        for i in range(len(Key_NameList)):
            self.index_dict[Key_NameList[i]] = i
        pass

    def addnew(self,x_line):
        self.x_line = x_line

        #如果当前行顺序小于写入行，快进
        while self.line_Compare() == 'LoadNext':
            self.line_LoadNext()

        #如果当前行完全被写入行覆盖，直接删除+快进
        while self.line_Compare() == 'DeleteDirectly':
            self.line_DeleteOrigin()
            self.line_LoadNext()

        if self.line_Compare() == 'AddTemp':
            self.temp_AddLine(self.line)
            return
        
        if self.line_Compare() == 'Calculate':
            self.line_DeleteOrigin()
            # line_Part_A 为前半段  /  line_Part_B 为后半段
            line_Part_A = []
            line_Part_B = []

            #如果前部有剩余，截取前半截直接写入
            if self.x_line[self.index_dict['起始年月']] > self.line[self.index_dict['起始年月']]:
                line_Part_A = self.line.copy
                line_Part_A[self.index_dict['终止年月']] = min(
                    line_Part_A[self.index_dict['终止年月']],
                    YM_Cal(self.x_line[self.index_dict['起始年月']] , -1 )
                ) #重新取一遍最小值，保险起见

                self.temp_AddLine(line_Part_A) #直接写入

            #如果后部有剩余,截取后半截保留
            if self.x_line[self.index_dict['终止年月']] < self.line[self.index_dict['终止年月']]:
                line_Part_B = self.line.copy
                line_Part_B[self.index_dict['起始年月']] = max(
                    line_Part_B[self.index_dict['起始年月']],
                    YM_Cal(self.x_line[self.index_dict['终止年月']] ,  1 )
                ) #重新取一遍最大值，保险起见

                self.line = line_Part_B #保留数据

        #计算完成后 递归调用，看看有没有下个情况，直到收到 'AddTemp' 为止
        self.addnew(x_line)


    def line_Compare(self):
        if self.line == []:  return 'AddTemp'

        if self.x_line[self.index_dict['员工字段ID']] > self.line[self.index_dict['员工字段ID']]: return 'LoadNext'
        if self.x_line[self.index_dict['员工字段ID']] < self.line[self.index_dict['员工字段ID']]: return 'AddTemp'
        
        #如果相邻且相等 则进行计算
        if self.x_line[self.index_dict['字段内容']] == self.line[self.index_dict['字段内容']]:
            if YM_Cal(self.x_line[self.index_dict['起始年月']] , -1 ) == self.line[self.index_dict['起始年月']]: return 'Calculate'
            if YM_Cal(self.x_line[self.index_dict['终止年月']] ,  1 ) == self.line[self.index_dict['终止年月']]: return 'Calculate'

        if self.x_line[self.index_dict['起始年月']] > self.line[self.index_dict['终止年月']]: return 'LoadNext'
        if self.x_line[self.index_dict['终止年月']] < self.line[self.index_dict['起始年月']]: return 'AddTemp'

        #如果当前行完全被写入行覆盖，直接删除换下一行
        if self.x_line[self.index_dict['起始年月']] <= self.line[self.index_dict['起始年月']] \
            and self.x_line[self.index_dict['终止年月']] >= self.line[self.index_dict['终止年月']]:

            return 'DeleteDirectly'

        return 'Calculate'


    def line_DeleteOrigin(self):
        
        if self.line_Edited == True: return
        
        self.ado_rs.Delete()
        self.line_Edited = True
    
    def line_LoadNext(self , LoadFirstLine = False):

        #如果该行处于编辑状态，换下一行之前写入当前行
        if self.line_Edited == True:
            self.temp_AddLine(self.line)
        
        #line初始化[]
        self.line = []

        if LoadFirstLine == True: self.ado_rs.MoveFirst()
        if self.ado_rs.EOF == True: return  #EOF则跳出
        if LoadFirstLine == False: self.ado_rs.MoveNext()

        for i in self.Key_NameList:
            self.line.append(self.ado_rs.Fields.Item(i).Value)  #使用ado.field.Item(x).Value进行读取

        self.line_Edited = False
    
    def temp_AddLine(self,add_line):

        #如果符合条件就合并到最后一行里
        #首先得至少有一行吧
        if self.temp_LinesToWrite != [] :
            #检测最后一行ID相等
            if self.temp_LinesToWrite[-1][self.index_dict['员工字段ID']] == add_line[self.index_dict['员工字段ID']]:
                #检测最后一行是否相等
                if self.temp_LinesToWrite[-1][self.index_dict['字段内容']] == add_line[self.index_dict['字段内容']]:
                    #检测是否相邻
                    if YM_Cal(self.temp_LinesToWrite[-1][self.index_dict['终止年月']] , 1 ) == add_line[self.index_dict['起始年月']]:
                        #如果都满足  直接延长最后一行
                        self.temp_LinesToWrite[-1][self.index_dict['终止年月']] = add_line[self.index_dict['终止年月']]

        self.temp_LinesToWrite.append(add_line)

    def temp_FinishAndWrite(self):

        #关闭后重新打开连接，省得麻烦
        self.ado_rs.Close()

        self.ado_rs.ActiveConnection = self.ado_con
        self.ado_rs.Source = self.table_name
        self.ado_rs.CursorType = 2
        self.ado_rs.LockType = 3
        self.ado_rs.Open()

        for i_line in self.temp_LinesToWrite :
            self.ado_rs.AddNew(self.Key_NameList , i_line)
            self.ado_rs.Update()
        
        self.ado_rs.Close() #结束，关闭


#直接定义一个类用来保存历史的 员工字段ID
class ID_Recorder():

    def __init__(self,Key_NameList):
        self.dict_员工ID = {'':0} # 初始带个零
        self.dict_字段ID = {'':0} # 初始带个零

        self.loc_员工编号 = Key_NameList.index('员工编号')
        self.loc_字段名称 = Key_NameList.index('字段名称')
        self.loc_员工字段ID = Key_NameList.index('员工字段ID')

    #convert 对 x_line 直接操作，没有返回值
    def convert(self , x_line):
        #先把读取的 员工字段ID 转换一下
        try:
            x_line[self.loc_员工字段ID] = int(x_line[self.loc_员工字段ID])
        except:
            x_line[self.loc_员工字段ID] = 0
        
        #比较计算 员工编号ID
        if x_line[self.loc_员工编号] in self.dict_员工ID :
            if x_line[self.loc_员工字段ID] > 0:
                if self.dict_员工ID[x_line[self.loc_员工编号]] != x_line[self.loc_员工字段ID] // 10000 :
                    print('ID 不一致警告 <员工编号>{} : <new>{} != <ori>{}').format(
                        x_line[self.loc_员工编号],
                        x_line[self.loc_员工字段ID] // 10000 , 
                        self.dict_员工ID[x_line[self.loc_员工编号]]
                    )
        else:
            if x_line[self.loc_员工字段ID] > 0: # 如果设定不冲突就进行设定
                self.dict_员工ID[x_line[self.loc_员工编号]] = x_line[self.loc_员工字段ID] // 10000
            else:
                self.dict_员工ID[x_line[self.loc_员工编号]] = max(self.dict_员工ID.items()) + 1

        #比较计算 字段名称ID
        if x_line[self.loc_字段名称] in self.dict_字段ID :
            if x_line[self.loc_员工字段ID] > 0:
                if self.dict_字段ID[x_line[self.loc_字段名称]] != x_line[self.loc_员工字段ID] % 10000 :
                    print('ID 不一致警告 <字段名称>{} : <new>{} != <ori>{}').format(
                        x_line[self.loc_字段名称],
                        x_line[self.loc_员工字段ID] % 10000 , 
                        self.dict_字段ID[x_line[self.loc_字段名称]]
                    )
        else:
            if x_line[self.loc_员工字段ID] > 0: # 如果设定不冲突就进行设定
                self.dict_字段ID[x_line[self.loc_字段名称]] = x_line[self.loc_员工字段ID] % 10000
            else:
                self.dict_字段ID[x_line[self.loc_字段名称]] = max(self.dict_字段ID.items()) + 1

        #最后覆盖掉 员工字段ID
        x_line[self.loc_员工字段ID] = self.dict_员工ID[x_line[self.loc_员工编号]] * 10000 + self.dict_字段ID[x_line[self.loc_字段名称]] 


class DataSet_Overwrite():

    def __init__(self , ado_con , table_name , Key_NameList , ID_Recorder):

        self.ado_con = ado_con
        self.table_name = table_name
        self.Key_NameList = Key_NameList
        self.ID_Recorder = ID_Recorder

        str_Keys = ' , '.join(Key_NameList)
        str_SQL = 'select {} from {} order by 员工字段ID , 起始年月'.format(str_Keys,self.table_name)

        self.ado_rs = win32com.client.Dispatch(r'ADODB.Recordset')
        self.ado_rs.ActiveConnection = self.ado_con
        self.ado_rs.Source = str_SQL
        self.ado_rs.CursorType = 2
        self.ado_rs.LockType = 3
        self.ado_rs.Open()

        self.lines = []
        self.ado_rs.MoveFirst()
        while self.ado_rs.EOF == False:
            read_line=[]
            for i in self.Key_NameList:
                read_line.append(self.ado_rs.Fields.Item(i).Value)  #使用ado.field.Item(x).Value进行读取

            self.ID_Recorder.convert(read_line) #读取转换 员工字段ID

            self.lines.append(read_line)
            self.ado_rs.MoveNext()

        # 开始排序，目标字段整合成元组之后再排序
        self.lines.sort(
            key = lambda x : (
                x[self.Key_NameList.index('员工字段ID')],
                x[self.Key_NameList.index('起始年月')],
                x[self.Key_NameList.index('终止年月')]
            ),
            reverse = False
        )

def main(dict_parameter):
    """用于在人员类别演算的数据库中合并各个时间段的标注信息
需要在运行前设置的 dict_parameter ：

<必填>
FileName : 需要处理的数据库文件名及路径

<选填>
Table_Origin : 需要进行覆盖的原始表格名称 (默认 = '标注数据整合')
Table_Overwrite : 用于对Table_Origin进行覆盖的新增数据 (默认 = '标注数据追加')

<生成>
ID_Recorder : 用来保存已经计算过的 员工字段ID"""

    #变量自检及初始化
    if not('FileName' in dict_parameter) :
        return 'Error 变量未设置 FileName : 需要处理的数据库文件名及路径'
    if not('Table_Origin' in dict_parameter) :
        dict_parameter['Table_Origin'] = '标注数据整合'
    if not('Table_Overwrite' in dict_parameter) :
        dict_parameter['Table_Overwrite'] = '标注数据追加'

    #命令：Connect
    #建立数据库 Connection

    #建立 ID_Recorder

    #命令：Calculate
    #链接创建 DataSet_Overwrite

    #链接创建 DataSet_Origin

    #叠加计算

    