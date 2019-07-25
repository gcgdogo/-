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
        self.ado_rs.LockType = 2
        self.ado_rs.CursorLocation = 2
        self.ado_rs.Open()

        #读取第一行
        self.line = [] #空行用空列表[]表示
        self.line_LoadNext(LoadFirstLine = True)

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

        #接着写
        
        #计算完成后 递归调用，看看有没有下个情况，直到收到 'AddTemp' 为止


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
        pass
    
    def line_LoadNext(self , LoadFirstLine = False):
        #line初始化[]
        self.line = []

        if LoadFirstLine == True: self.ado_rs.MoveFirst()
        if self.ado_rs.EOF == True: return  #EOF则跳出
        if LoadFirstLine == False: self.ado_rs.MoveNext()

        for i in self.Key_NameList:
            self.line.append(self.ado_rs.Fields.Item(i).Value)  #使用ado.field.Item(x).Value进行读取

    
    def temp_AddLine(self,x_line):
        pass

    def temp_FinishAndWrite(self):
        pass