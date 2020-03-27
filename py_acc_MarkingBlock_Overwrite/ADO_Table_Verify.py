import win32com.client
class Table_Verify:
    """
    验算SQL语句：
    vb版："SELECT 
        format(count(ID) mod 1000000 ,""000000"") & ""-"" & 
        format(sum(ID mod 983) mod 1000000 ,""000000"") & ""-"" & 
        format(sum(ID mod 997) mod 1000000 ,""000000"") As Check_Code FROM T1;"

    为保障稳定性，牺牲运算速度做了两个余数的求和，效率不高
    Access到Python的数据似乎可以通过重新连接数据库来保障数据一致性，Access无法断开数据库，需要添加校验码

    根据新增项ID比旧项大的规律，需要新增项与删除项数目一致，且平均值的差值正好为 983*997 才能漏过
    """
    def __init__(self,ADO_Connection,x_TableName,x_ColumnName = "ID"):
        #初始设置
        self.ADO_Connection = ADO_Connection

        #用来统计重新验算次数
        self.Retry_Times = 0
        self.SQLstring =""
        self.set_TargetTable(x_TableName,x_ColumnName)

    def set_TargetTable(self,x_TableName,x_ColumnName = "ID"):
        self.SQLstring = """SELECT 
        format(count({1}) mod 1000000 ,"000000") & "-" & 
        format(sum({1} mod 983) mod 1000000 ,"000000") & "-" & 
        format(sum({1} mod 997) mod 1000000 ,"000000") As Check_Code 
        FROM {0};""".format(x_TableName,x_ColumnName)

    def cal_CheckCode(self):
        #进行Recordset连接
        ado_rs = win32com.client.Dispatch(r'ADODB.Recordset')
        ado_rs.ActiveConnection = self.ADO_Connection
        ado_rs.Source = self.SQLstring

        #print(ado_rs.Source)

        ado_rs.CursorType = 3 #静态游标
        ado_rs.LockType = 1 #ReadOnly
        ado_rs.Open()
        #提取结果
        str_CheckCode = ado_rs.Fields.Item(0).Value
        #关闭连接
        ado_rs.Close()

        return str_CheckCode

