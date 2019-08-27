Option Compare Database
Option Explicit

'需要引用 Microsoft WinHttp Services

Dim TT_PortSeed As Integer, TT_ServerName As String
Dim TT_URL as String

Function TT_Server_Start(In_PythonFileName As String, In_PortSeed As Integer, In_ServerName As String)
    Dim CommandString As String
    
    CommandString = "python " + In_PythonFileName + " --Server/Name=" + In_ServerName + " --Server/PortSeed=" + CStr(In_PortSeed)
    Diary_Add "Running", "TT_Server_Start : " + CommandString
    Shell CommandString  '执行命令启动程序
    
    '变量设置
    TT_PortSeed = In_PortSeed
    TT_ServerName = In_ServerName

    Call Get_URL()
    Call TT_Server_Initalize()

    Diary_Add "Message", "TT 初始化完毕 : URL = " + TT_URL
End Function

Private Function TT_Server_Initalize() 
'初始化
    Request_Post "/Parameter/FileName", "'" + Application.CurrentProject.FullName + "'"
    Request_Post "/Parameter/Table_Origin","'标注数据整合_python'"
    Request_Post "/Parameter/Table_Overwrite","'标注数据追加'"

    Request_Get "/Execute/Connect" '链接数据库
End Function

Private Function SourceTable_RecordCount()
    Dim ADO_rs as New ADODB.Recordset
    
    ADO_rs.ActiveConnection = CurrentProject.Connection
    ADO_rs.Source = "标注数据追加"
    ADO_rs.CursorType = adOpenStatic
    ADO_rs.LockType = adLockReadOnly
    ADO_rs.Open
    SourceTable_RecordCount = ADO_rs.RecordCount
    
    ADO_rs.Close
End Function

Function TT_Server_Calculate()
    Dim Start_Time as Double
    Dim RecordCount as Long
    Dim Server_ReturnString as String
    Dim Pardon_Count as Integer

    RecordCount = SourceTable_RecordCount()  '记录一下数据量，要不然可能有部分数据没写完，调用一下Access的数据连接，看看数据写完没

    Start_Time = Timer()
    Diary_Add "Running", "TT_Execute[" & RecordCount & "]" '提示开始运行

    Server_ReturnString = Request_Get("/Execute/Calculate") '运行命令

    '如果没有得到运行结果，执行PardonMe
    Pardon_Count = 0
    do while Server_ReturnString = "[Request_Get_Fail]"

        Pardon_Count = Pardon_Count + 1
        Server_ReturnString = Request_Get("/PardonMe") '重新获取

        Diary_Add "Running", "TT_Execute[" & RecordCount & "] Pardon_Count = " & Pardon_Count '提示Pardon中
        '如果执行半天了都没有回复：报错吧！！！ 初步设置为 300s
        if (Timer()-Start_Time) > 300 then msgbox(0/0/0/0/0/0/0)  '报错吧！！！！！
    loop

    Diary_Add "Message", "TT_Execute[" & RecordCount & "]:" & int((Timer()-Start_Time)*1000) & "ms " & Server_ReturnString
    
    if RecordCount <> SourceTable_RecordCount() then
        msgbox("发现错误!  执行计算后发现有新增的记录 " & RecordCount & " > " & SourceTable_RecordCount())
        msgbox(0/0/0/0/0/0/0)  '报错吧！！！！！
    end if
End Function

Private Function Get_URL()
    Dim Start_Time as Double
    Dim Max_I as Integer ,I as Integer , Get_Port as Integer
    Dim Got_ServerName as String
    Start_Time = timer()
    for Max_I=0 to 5  '0,0,1,0,1,2,3的循序循环
        for I=0 to Max_I
            DoEvents
            TT_URL = "http://localhost:" + cstr(TT_PortSeed + I * 567)
            
            Got_ServerName = Request_Get("/Parameter/Server/Name")

            Diary_Add "Running", "Try_Connection : " + TT_URL + " = " + Got_ServerName

            if Got_ServerName = TT_ServerName then Exit Function
        next
    next
    msgbox("无法获取python程序的URL")
    Get_URL = 0/0/0/0/0/0/0/0  '直接报错好啦
End Function

Function Request_Get(URL_append as String)
    On Error Goto Request_Get_Fail
    Static WinReq As WinHttpRequest
    Static WinReq_URL As String
    if WinReq_URL = "" then set WinReq = New WinHttpRequest
    if WinReq_URL <> TT_URL + URL_append then
        WinReq.Open "GET", TT_URL + URL_append 'URL变化就重新OPEN一下
        WinReq_URL = TT_URL + URL_append
    end if
    WinReq.Send
    Request_Get = WinReq.ResponseText
    Exit Function
Request_Get_Fail:
    Request_Get = "[Request_Get_Fail]"
End Function

Function Request_Post(URL_append as String , str_Message as String)
    On Error Goto Request_Post_Fail
    Static WinReq As WinHttpRequest
    Static WinReq_URL As String
    if WinReq_URL = "" then set WinReq = New WinHttpRequest
    if WinReq_URL <> TT_URL + URL_append then
        WinReq.Open "POST", TT_URL + URL_append 'URL变化就重新OPEN一下
        WinReq_URL = TT_URL + URL_append
    end if
    WinReq.Send str_Message
    Request_Post = WinReq.ResponseText
    Exit Function
Request_Post_Fail:
    Request_Post = "[Request_Post_Fail]"
End Function

sub TT_Test()
    TT_Server_Start "F:\Site\GitHub\-\py_acc_MarkingBlock_Overwrite\MarkingBlock_Overwrite.py",8460,"MarkingBlock_Overwrite"
End sub

Sub TT_Count()
    MsgBox (SourceTable_RecordCount())
End Sub