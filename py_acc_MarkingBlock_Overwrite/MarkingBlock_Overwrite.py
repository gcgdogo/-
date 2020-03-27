import win32com.client
import HTTP_Thread_Trigger
import numpy
#import pickle # 测试用

import ADO_Table_Verify #用于校验数据库

def YM_Cal( YM_input , X_offset ):
    Months = (YM_input // 100) * 12 + (YM_input % 100)
    Months = Months + X_offset - 1
    return (Months // 12) * 100 + (Months % 12) + 1

class DataSet_Origin:

    def __init__(self,ado_con,table_name,Key_NameList):
        #需要写入的行先临时保存起来
        self.temp_LinesToWrite = []

        #保存运行记录
        self.list_log = []
        for i in range(10):
            self.list_log.append(('',[],[]))
        #保存ado链接
        self.ado_con = ado_con

        #保存表名
        self.table_name = table_name

        self.table_delete_name = 'MarkingBlock_待删除行号'

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
        self.ado_rs.CursorType = 3 #静态游标
        self.ado_rs.LockType = 1 #只读
        self.ado_rs.Open()

        try:
            self.ado_rs.MoveFirst()
            self.Ori_lines = numpy.array(self.ado_rs.GetRows()).T.tolist()  #直接一次性读取，用numpy进行转置和计算
        except:
            self.Ori_lines = []
        self.Ori_lines_finger = 0
        self.ado_rs.Close()  #读取完就直接关闭吧

        #Delete之后无法直接接 update 且效率慢，改为存储行号到单独的表里统一删除
        self.ado_con.Execute(
            'DELETE * FROM {0};'.format(self.table_delete_name)
        )  #先把待删除列表清空

        self.ado_rs_delete = win32com.client.Dispatch(r'ADODB.Recordset')
        self.ado_rs_delete.ActiveConnection = self.ado_con
        self.ado_rs_delete.Source = 'select {} from {}'.format('ID',self.table_delete_name)
        self.ado_rs_delete.CursorType = 2
        self.ado_rs_delete.LockType = 3
        self.ado_rs_delete.Open()


        #读取第一行
        self.line = [] #空行用空列表[]表示
        self.line_NeedToWrite = False
        self.line_Deleted = False

        self.line_LoadNext(LoadFirstLine = True)

        #保存顺序位置
        self.index_dict = {}
        for i in range(len(Key_NameList)):
            self.index_dict[Key_NameList[i]] = i
        pass
        
        #记录执行情况
        self.count_Input = 0
        self.count_Delete = 0
        self.count_AddTemp = 0
        self.count_Write = 0

    def addnew(self,x_line):
        #print('x_line' + repr(x_line) + repr(self.line))
        self.x_line = x_line
        self.count_Input = self.count_Input + 1

        #如果当前行顺序小于写入行，快进
        while self.line_Compare() == 'LoadNext':
            self.line_LoadNext()

        #如果当前行完全被写入行覆盖，直接删除+快进
        while self.line_Compare() == 'DeleteDirectly':
            self.line_DeleteOrigin()
            self.line_NeedToWrite = False  ######得在删除之后把编辑标记解除掉，要不然换下一行的时候又给加进去了
            self.line_LoadNext()

        if self.line_Compare() == 'AddTemp':

            self.list_log = self.list_log[1:]
            self.list_log.append(('AddTemp',self.line.copy(),self.x_line.copy()))

            self.temp_AddLine(self.x_line)
            return
        
        if self.line_Compare() == 'Calculate':

            self.list_log = self.list_log[1:]
            self.list_log.append(('Calculate',self.line.copy(),self.x_line.copy()))

            self.line_DeleteOrigin()
            # line_Part_A 为前半段  /  line_Part_B 为后半段
            line_Part_A = []
            line_Part_B = []

            #如果前部有剩余，截取前半截直接写入
            if self.x_line[self.index_dict['起始年月']] > self.line[self.index_dict['起始年月']]:
                line_Part_A = self.line.copy()
                line_Part_A[self.index_dict['终止年月']] = min(
                    line_Part_A[self.index_dict['终止年月']],
                    YM_Cal(self.x_line[self.index_dict['起始年月']] , -1 )
                ) #重新取一遍最小值，保险起见

                self.temp_AddLine(line_Part_A) #直接写入
                self.line_NeedToWrite = False  ######把编辑标记解除掉，要不然换下一行的时候又给加进去了

            #如果后部有剩余,截取后半截保留
            if self.x_line[self.index_dict['终止年月']] < self.line[self.index_dict['终止年月']]:
                line_Part_B = self.line.copy()
                line_Part_B[self.index_dict['起始年月']] = max(
                    line_Part_B[self.index_dict['起始年月']],
                    YM_Cal(self.x_line[self.index_dict['终止年月']] ,  1 )
                ) #重新取一遍最大值，保险起见

                self.line = line_Part_B #保留数据

                self.temp_AddLine(self.x_line) #后面都有剩余了，这行可以写进去了
                self.line_NeedToWrite = True  ######加上标记，需要进行写入
                return
            else:
                self.line_LoadNext()  #后面要是没有节余就可以移动到下一行了


        #计算完成后 递归调用，看看有没有下个情况，直到收到 'AddTemp' 或 'Calculate且后面有节余'为止
        self.addnew(x_line)


    def line_Compare(self):
        if self.line == []:  return 'AddTemp'

        if self.x_line[self.index_dict['员工字段ID']] > self.line[self.index_dict['员工字段ID']]: return 'LoadNext'
        if self.x_line[self.index_dict['员工字段ID']] < self.line[self.index_dict['员工字段ID']]: return 'AddTemp'
        
        #如果字段名称或员工编号不相等，赶紧报错！！
        if self.x_line[self.index_dict['字段名称']] != self.line[self.index_dict['字段名称']] \
            or self.x_line[self.index_dict['员工编号']] != self.line[self.index_dict['员工编号']]:

            print('字段名称不相等警告！')
            print('已有:{}'.format(repr(self.line)))
            print('写入:{}'.format(repr(self.x_line)))
            print(0/0/0/0/0/0/0/0/0/0)   ############报错吧！！！

        #如果相邻且相等 则进行计算
        if self.x_line[self.index_dict['字段内容']] == self.line[self.index_dict['字段内容']]:
            if YM_Cal(self.x_line[self.index_dict['起始年月']] , -1 ) == self.line[self.index_dict['终止年月']]: return 'Calculate'
            if YM_Cal(self.x_line[self.index_dict['终止年月']] ,  1 ) == self.line[self.index_dict['起始年月']]: return 'Calculate'

        if self.x_line[self.index_dict['起始年月']] > self.line[self.index_dict['终止年月']]: return 'LoadNext'
        if self.x_line[self.index_dict['终止年月']] < self.line[self.index_dict['起始年月']]: return 'AddTemp'

        #如果当前行完全被写入行覆盖，直接删除换下一行
        if self.x_line[self.index_dict['起始年月']] <= self.line[self.index_dict['起始年月']] \
            and self.x_line[self.index_dict['终止年月']] >= self.line[self.index_dict['终止年月']]:

            return 'DeleteDirectly'

        return 'Calculate'


    def line_DeleteOrigin(self):
        
        if self.line_Deleted == True: return
        
        self.ado_rs_delete.AddNew(['ID'],self.line[0:1]) #写入第一个值，就是ID

        #self.ado_rs.Delete()
        self.line_Deleted = True
        self.line_NeedToWrite = True
        self.count_Delete = self.count_Delete + 1
    
    def line_LoadNext(self , LoadFirstLine = False):

        #如果该行处于编辑状态，换下一行之前写入当前行
        if self.line_NeedToWrite == True:
            self.temp_AddLine(self.line)
        
        if self.Ori_lines_finger < len(self.Ori_lines):  #直接从已读取的各行中截取一行就行
            self.line = self.Ori_lines[self.Ori_lines_finger]
            self.Ori_lines_finger = self.Ori_lines_finger + 1
        else:
            self.line = []

        self.line_Deleted = False  #恢复编辑状态
        self.line_NeedToWrite = False

    def temp_AddLine(self,add_line):

        self.count_AddTemp = self.count_AddTemp + 1

        #如果符合条件就合并到最后一行里
        #首先得至少有一行吧
        if len(self.temp_LinesToWrite) > 0 :
            #print(self.temp_LinesToWrite[-1])
            #检测最后一行ID相等
            if self.temp_LinesToWrite[-1][self.index_dict['员工字段ID']] == add_line[self.index_dict['员工字段ID']]:
                #如果后加的年月小于前一行，赶紧报错啊！！！
                if self.temp_LinesToWrite[-1][self.index_dict['终止年月']] >= add_line[self.index_dict['起始年月']]:
                    print("后加的年月小于前一行，赶紧报错啊！！！")
                    print(self.temp_LinesToWrite[-1])
                    print(add_line)

                    print("self.list_log:")
                    for i in self.list_log:
                        print(i)

                    print(0/0/0/0/0/0/0/0/0/0/0/0) #直接报错就得了

                #检测最后一行是否相等
                if self.temp_LinesToWrite[-1][self.index_dict['字段内容']] == add_line[self.index_dict['字段内容']]:
                    #检测是否相邻
                    if YM_Cal(self.temp_LinesToWrite[-1][self.index_dict['终止年月']] , 1 ) == add_line[self.index_dict['起始年月']]:
                        #如果都满足  直接延长最后一行
                        self.temp_LinesToWrite[-1][self.index_dict['终止年月']] = add_line[self.index_dict['终止年月']]
                        return #直接返回

        self.temp_LinesToWrite.append(add_line)

    def temp_FinishAndWrite(self):
        
        #先尝试把位置移动到下一行，避免最后一行写入后，当前line处于编辑状态
        self.line_LoadNext()

        #一次性删除
        self.ado_con.Execute(
            'DELETE {0}.* FROM {0} INNER JOIN {1} ON {0}.ID = {1}.ID;'.format(self.table_name,self.table_delete_name)
        )

        self.ado_rs.ActiveConnection = self.ado_con
        self.ado_rs.Source = self.table_name
        self.ado_rs.CursorType = 2
        self.ado_rs.LockType = 3
        self.ado_rs.Open()

        self.count_Write = len(self.temp_LinesToWrite)

        for i_line in self.temp_LinesToWrite :

            #监视符合条件的行是否写入了
            #if i_line[self.index_dict['员工编号']] == '1301931' : print('write{}'.format(repr(i_line)))

            self.ado_rs.AddNew(self.Key_NameList[1:] , i_line[1:]) #第一个位置是行号，不参与导入
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
            #print( x_line)
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
                self.dict_员工ID[x_line[self.loc_员工编号]] = max(self.dict_员工ID.values()) + 1

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
                self.dict_字段ID[x_line[self.loc_字段名称]] = max(self.dict_字段ID.values()) + 1

        #最后覆盖掉 员工字段ID
        x_line[self.loc_员工字段ID] = self.dict_员工ID[x_line[self.loc_员工编号]] * 10000 + self.dict_字段ID[x_line[self.loc_字段名称]] 


class DataSet_Overwrite():

    def __init__(self , ado_con , table_name , Key_NameList , ID_Recorder):

        #print('DataSet_Overwrite 初始化')
        self.ado_con = ado_con
        self.table_name = table_name
        self.Key_NameList = Key_NameList
        self.ID_Recorder = ID_Recorder

        str_Keys = ' , '.join(Key_NameList)
        str_SQL = 'select {} from {} order by 员工字段ID , 起始年月'.format(str_Keys,self.table_name)

        self.ado_rs = win32com.client.Dispatch(r'ADODB.Recordset')
        self.ado_rs.ActiveConnection = self.ado_con
        self.ado_rs.Source = str_SQL
        self.ado_rs.CursorType = 3 #静态游标
        self.ado_rs.LockType = 1 #ReadOnly
        self.ado_rs.Open()
        #print('DataSet_Overwrite.lines 读取')

        print("DataSet_Overwrite.ado_rs.RecordCount = {}".format(self.ado_rs.RecordCount))

        try:
            self.ado_rs.MoveFirst()
            self.lines = numpy.array(self.ado_rs.GetRows()).T.tolist()  #直接一次性读取，用numpy进行转置和计算
        except:
            self.lines = []

        # if len(self.lines)!=self.ado_rs.RecordCount :
        #     print("读取行数不一致！: {} _ {}".format(len(self.lines),self.ado_rs.RecordCount))
        #     pickle.dump(
        #         self.lines,
        #         open(r'F:\Database\Development\人员类别演算\py_HTTP_ThreadTriger_改造测试\Error_Data_{}_{}.txt'.format(
        #             len(self.lines),
        #             self.ado_rs.RecordCount
        #             ), 'wb')
        #     ) 


        #print(self.lines)
        for read_line in self.lines:
            #监视符合条件的行是否尝试进行转换
            #if read_line[1] == '1301931' : print('convert{}'.format(repr(read_line)))

            self.ID_Recorder.convert(read_line)
        # while self.ado_rs.EOF == False:
        #     read_line=[]
        #     for i in self.Key_NameList:
        #         read_line.append(self.ado_rs.Fields.Item(i).Value)  #使用ado.field.Item(x).Value进行读取

        #     self.ID_Recorder.convert(read_line) #读取转换 员工字段ID

        #     self.lines.append(read_line)
        #     self.ado_rs.MoveNext()


        #print('DataSet_Overwrite.lines 排序')
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
操作命令[Thread/str_Path]：
    Connect
    Calculate
需要在运行前设置的 dict_parameter ：
<必填>
FileName : 需要处理的数据库文件名及路径

<选填>
Table_Origin : 需要进行覆盖的原始表格名称 (默认 = '标注数据整合')
Table_Overwrite : 用于对Table_Origin进行覆盖的新增数据 (默认 = '标注数据追加')
Key_NameList : 需要进行组合的字段列表(默认 =  ['ID' , '员工编号' , '姓名' , '起始年月' , '终止年月' , '字段名称' , '字段内容' , '员工字段ID']) 第一个用于存储行号，不参与导入

<生成>
ADO_Connection : 用来保存数据库连接
ID_Recorder : 用来保存已经计算过的 员工字段ID"""

    #变量自检及初始化
    if not('FileName' in dict_parameter) :
        return 'Error 变量未设置 FileName : 需要处理的数据库文件名及路径'
    if not('Table_Origin' in dict_parameter) :
        dict_parameter['Table_Origin'] = '标注数据整合'
    if not('Table_Overwrite' in dict_parameter) :
        dict_parameter['Table_Overwrite'] = '标注数据追加'
    if not('Key_NameList' in dict_parameter) :
        dict_parameter['Key_NameList'] = ['ID' , '员工编号' , '姓名' , '起始年月' , '终止年月' , '字段名称' , '字段内容' , '员工字段ID']

    if not('Thread/str_Path' in dict_parameter):
        print('无效调用，直接结束')
        return

    win32com.client.pythoncom.CoInitialize() #看看这么能不能解决多线程的问题

    #默认返回字符串
    Return_String = '未找到对应命令，直接结束'

    #命令：Connect
    if dict_parameter['Thread/str_Path'] == 'Connect' :
        #虽然这个Connection没啥用，但是如果不保持一个连接，保持占用状态，每次重连都会很慢
        #建立数据库 Connection
        dict_parameter['ADO_Connection'] = win32com.client.Dispatch(r'ADODB.Connection')
        dict_parameter['ADO_Connection'].Open(r'Provider=Microsoft.ACE.OLEDB.12.0;Data Source={}'.format(dict_parameter['FileName']))

        #建立 ID_Recorder
        dict_parameter['ID_Recorder'] = ID_Recorder(dict_parameter['Key_NameList'])
 
        Return_String = '成功连接数据库文件[{}]'.format(dict_parameter['FileName'])

    #命令：Calculate
    if dict_parameter['Thread/str_Path'] == 'Calculate' :

        #测试代码：看看每次都重新连接数据库行不行
        win32_ADO_Connection = win32com.client.Dispatch(r'ADODB.Connection')
        win32_ADO_Connection.Open(r'Provider=Microsoft.ACE.OLEDB.12.0;Data Source={}'.format(dict_parameter['FileName']))

    #链接创建 DataSet_Overwrite
        DS_Overwrite = DataSet_Overwrite(
            ado_con = win32_ADO_Connection,
            table_name = dict_parameter['Table_Overwrite'],
            Key_NameList = dict_parameter['Key_NameList'],
            ID_Recorder = dict_parameter['ID_Recorder']
        )
    #链接创建 DataSet_Origin
        DS_Origin = DataSet_Origin(
            ado_con = win32_ADO_Connection,
            table_name = dict_parameter['Table_Origin'],
            Key_NameList = dict_parameter['Key_NameList']
        )

    #创建校验码计算对象 Table_Verify
        DS_Table_Verify = ADO_Table_Verify.Table_Verify(
            ADO_Connection = win32_ADO_Connection,
            x_TableName = dict_parameter['Table_Origin'],
            x_ColumnName = "ID"
        )

    #叠加计算
        #遍历各行
        for i_line in DS_Overwrite.lines:
            #监视符合条件的行是否尝试添加了
            #if i_line[1] == '1301931' : print('try_add{}'.format(repr(i_line)))

            DS_Origin.addnew(i_line)
        #最后写入
        DS_Origin.temp_FinishAndWrite()

        Return_String = '数据整合完毕 Lines = {} Input = {} Delete = {} AddTemp = {} Write = {}'.format(
            len(DS_Overwrite.lines),
            DS_Origin.count_Input,
            DS_Origin.count_Delete,
            DS_Origin.count_AddTemp,
            DS_Origin.count_Write
        )

        # CheckCode 与返回信息 用 :: 分割
        Return_String = DS_Table_Verify.cal_CheckCode() + "::" + Return_String

        win32_ADO_Connection.Close()

    win32com.client.pythoncom.CoUninitialize() #释放多线程

    return Return_String

if __name__ == '__main__':

    HTTP_Thread_Trigger.dict_parameter['Thread/Source'] = main
    HTTP_Thread_Trigger.start_server()

