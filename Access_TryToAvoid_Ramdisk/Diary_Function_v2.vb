Option Compare Database
Option Explicit

Public Diary_Application as Application , Diary_HeadString as String

Sub Diary_Application_Set(X_App as Application)
    Set Diary_Application = X_App
End Sub

'  [Call:LangxinData][Call]
Sub Diary_HeadString_Set(X_String as String)
    Diary_HeadString = X_String
End Sub

Function Diary_Add(In_Type As String, In_Txt As String)

    Dim DiaryADO As ADODB.Recordset
    Set DiaryADO = New ADODB.Recordset
    Dim DiarySQL As String

    '如果已设置 Diary_Application 则反向调用 Diary_Add
    If Not (Diary_Application Is Nothing) Then
        Diary_Application.Run "Diary_Add", In_Type, Diary_HeadString & In_Txt
        Exit Function
    End If

    DiarySQL = "INSERT INTO Temp_Diary ( Type, [Time], ms, Txt ) SELECT """ & In_Type & """,now(),""" & CStr((Timer() * 1000) Mod 1000) & """ , """ & In_Txt & """;"
    DiaryADO.Source = DiarySQL
    DiaryADO.ActiveConnection = CurrentProject.Connection
    DiaryADO.Open
    
    Application.Forms("diary").Requery
    Application.Forms("diary").Repaint
    
    Do While Form_Diary.Dirty = True
        DoEvents
        DoEvents
        DoEvents
    Loop
    
End Function

'报告对应命令执行发生时间
Function Diary_TimeElapsed( Optional Condition_Exp as String = "" , Optional Exception_Exp as String = "[Call:*" ) as Variant
    Dim CommandTime as Variant , TimeElapsed as Double
    CommandTime = Diary_CommandTime( "*" & Condition_Exp & "*" , Exception_Exp )
    If CommandTime = "未找到对应记录" Then
        Diary_Add "Message" , "[TimeElapsed] 未找到对应记录"
        Exit Function
    End If

    TimeElapsed = ( Now() - CommandTime ) * 24 * 60 * 60
    If TimeElapsed > 7200 Then
        Diary_Add "Message" , "[TimeElapsed] 时间过长"
        Exit Function
    End If

    Diary_Add "Message" , "[TimeElapsed] = " & Format( TimeElapsed , "#,##0" ) & "s"
    
End Function

'SQL中 [ 不能正常用 like 进行匹配，需要用RegExp.Replace 进行替换,只有左中括号不能匹配，右中括号加[]会出错
Function Diary_CommandTime(Condition_Exp as String , Exception_Exp as String) as Variant
    Dim R_Exp as New RegExp
    Dim Condition_Exp_Edited as String , Exception_Exp_Edited as String
    Dim Get_CommandTime as Variant
    If Not (Diary_Application Is Nothing) Then
        Diary_CommandTime = Diary_Application.Run ("Diary_CommandTime", Diary_HeadString & Condition_Exp , Diary_HeadString & Exception_Exp)
        Exit Function
    End If

    R_Exp.IgnoreCase=True
    R_Exp.Global=True
    R_Exp.Pattern="(\[)"
    Condition_Exp_Edited = R_Exp.Replace(Condition_Exp,"[$1]")
    Exception_Exp_Edited = R_Exp.Replace(Exception_Exp,"[$1]")

    Get_CommandTime = DMax("Time","Temp_Diary","Type=""Command"" and Txt Like """ & Condition_Exp_Edited & """ and Txt Not Like """ & Exception_Exp_Edited & """")
    'inputbox Prompt := "DMax" , Default := "Type=""Command"" and Txt Like """ & Condition_Exp_Edited & """ and Txt Not Like """ & Exception_Exp_Edited & """"
    IF isnull(Get_CommandTime) then Get_CommandTime = "未找到对应记录"
    'msgbox(Get_CommandTime)
    Diary_CommandTime=Get_CommandTime
End Function

Sub Test()
    Diary_Add "asdf", "asdf"
End Sub