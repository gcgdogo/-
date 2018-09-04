Option Compare Database
Option Explicit

Sub Diary_Application_Set(X_App as Application)

End Sub

Function Diary_Application_Get() as Application

End Function


'  [Call:LangxinData][Call]
Function Diary_HeadString() as String

End Function

Function Diary_Add(In_Type As String, In_Txt As String)

    Dim DiaryADO As ADODB.Recordset
    Set DiaryADO = New ADODB.Recordset
    Dim DiarySQL As String
    DiarySQL = "INSERT INTO Temp_Diary ( Type, [Time], ms, Txt ) SELECT """ & In_Type & """,now(),""" & CStr((Timer() * 1000) Mod 1000) & """ , """ & In_Txt & """;"
    DiaryADO.Source = DiarySQL
    DiaryADO.ActiveConnection = CurrentProject.Connection
    DiaryADO.Open
    
    Application.Forms("diary").Requery
    Application.Forms("diary").Repaint
    
    Do While Form_Diary.Dirty = True
        DoEvents
    Loop
    
End Function

Sub Test()
x = Diary_Add_2("asdf", "asdf")
End Sub