Option Compare Database
Option Explicit

'修改 v2.1 : TimeElapsed改为以模块内置记录的 Command_Time 为基础
Dim Command_Time as Double

Public Diary_Application as Application , Diary_HeadString as String

Sub Diary_Application_Set(X_App as Application)
    Set Diary_Application = X_App
End Sub

'  [Call:LangxinData][Call]
Sub Diary_HeadString_Set(X_String as String)
    Diary_HeadString = X_String
End Sub

'为实现Command计时，将Diary_Add拆成两个部分

Function Diary_Add(In_Type As String, In_Txt As String)

    'Command 计时
    If In_Type="Command" Then
        Command_Time = Timer()
    End If
    
    Call Diary_WriteNewLine(In_Type, In_Txt)

End Function

'Diary_WriteNewLine 负责写入部分，不计时

Function Diary_WriteNewLine(In_Type As String, In_Txt As String)
    Dim DiaryADO As ADODB.Recordset
    Set DiaryADO = New ADODB.Recordset
    Dim DiarySQL As String

    Dim TargetForm_Name as String , I as Integer

    '如果已设置 Diary_Application 则反向调用 Diary_Add
    'v2.1 ： 改成调用 Diary_WriteNewLine 避免 Command 重复计时
    If Not (Diary_Application Is Nothing) Then
        Diary_Application.Run "Diary_WriteNewLine", In_Type, Diary_HeadString & In_Txt
        Exit Function
    End If


    '开始调用新增记录

    TargetForm_Name = "Terminal"

    For I = 0 to Forms.count - 1
        If Forms(I).name = TargetForm_Name Then Exit For
    Next

    If I >= Forms.count Then DoCmd.OpenForm TargetForm_Name

    Call Forms(TargetForm_Name).WriteNewLine(In_Type, In_Txt)


    ' DiarySQL = "INSERT INTO Temp_Diary ( Type, [Time], ms, Txt ) SELECT """ & In_Type & """,now(),""" & CStr((Timer() * 1000) Mod 1000) & """ , """ & In_Txt & """;"
    ' DiaryADO.Source = DiarySQL
    ' DiaryADO.ActiveConnection = CurrentProject.Connection
    ' DiaryADO.Open
    
    ' Application.Forms("diary").Requery
    ' Application.Forms("diary").Repaint
    
    ' Do While Form_Diary.Dirty = True
    '     DoEvents
    '     DoEvents
    '     DoEvents
    ' Loop
    
End Function

'报告对应命令执行发生时间

Function Diary_TimeElapsed() as Variant
    If Command_Time=0 Then
        Diary_Add "Message" , "[TimeElapsed] 未找到对应记录"
    Else
        Diary_Add "Message" , "[TimeElapsed] = " & Format( (Timer() - Command_Time) , "#,##0.00" ) & "s"
        Command_Time = 0
    End If
End Function

Sub Test()
    Diary_Add "asdf", "asdf"
End Sub