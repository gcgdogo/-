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

Sub Test()
    Diary_Add "asdf", "asdf"
End Sub