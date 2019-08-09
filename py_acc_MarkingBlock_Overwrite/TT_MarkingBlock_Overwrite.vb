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
    Request_Post "/Parameter/FileName",Application.CurrentProject.FullName
    Request_Post "/Parameter/Table_Origin","标注数据整合_python"
    Request_Post "/Parameter/Table_Overwrite","标注数据追加"

    Request_Get "/Execute/Connect" '链接数据库
End Function

Function TT_Server_Calculate()
    Request_Get "/Execute/Calculate" '运行命令
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
    if WinReq_URL <> TT_URL then WinReq.Open "GET", TT_URL + URL_append  '直接调用已经设置好的URL
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
    if WinReq_URL <> TT_URL then WinReq.Open "POST", TT_URL + URL_append  '直接调用已经设置好的URL
    WinReq.Send str_Message
    Request_Post = WinReq.ResponseText
    Exit Function
Request_Post_Fail:
    Request_Post = "[Request_Post_Fail]"
End Function

sub TT_Test()
    TT_Server_Start "F:\Site\GitHub\-\py_acc_MarkingBlock_Overwrite\MarkingBlock_Overwrite.py",8460,"MarkingBlock_Overwrite"
End sub