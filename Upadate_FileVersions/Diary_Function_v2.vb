Option Compare Database
Option Explicit

Static Diary_Application as Application , Diary_HeadString as String

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

    '检测 Diary_Application 是否已设定 (Nothing)
    If Diary_Application is Nothing Then Set Diary_Application = Application

    DiarySQL = "INSERT INTO Temp_Diary ( Type, [Time], ms, Txt ) SELECT """ & In_Type & """,now(),""" & CStr((Timer() * 1000) Mod 1000) & """ , """ & In_Txt & """;"
    DiaryADO.Source = DiarySQL
    DiaryADO.ActiveConnection = Diary_Application.CurrentProject.Connection
    DiaryADO.Open
    
    Diary_Application.Forms("diary").Requery
    Diary_Application.Forms("diary").Repaint
    
    Do While Diary_Application.Forms("diary").Dirty = True
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Loop
    
End Function

Sub Test()
x = Diary_Add_2("asdf", "asdf")
End Sub