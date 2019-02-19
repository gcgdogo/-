Option Explicit

Dim X_Dict As New Dictionary
Dim TimerFlag As Double
Dim Last_DictSource As String
Dim R_Exp As New RegExp

Public Function Get_WordForm(X_Word As String, Dict_Source As String)
    
    '如果距离上次执行超过规定时间就重置词典
    If Timer() - TimerFlag > 10 Or Last_DictSource <> Dict_Source Then Call Get_WordForm_LoadDict(Dict_Source)

    Get_WordForm = Try_WordForm(X_Word, "", "Ori")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "", "LCase")

    '加s系列
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "s$", "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "es$", "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ies$", "y")

    '加ing系列
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ing$", "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ing$", "e")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ying$", "ie")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "(.)\1ing$", "$1")

    '加ed系列
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ed$", "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ed$", "e")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "ied$", "y")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word, "(.)\1ed$", "$1")

    TimerFlag = Timer()

End Function


Private Sub Get_WordForm_LoadDict(Dict_Source As String)
    Dim ADO_rs As New ADODB.Recordset
    ADO_rs.ActiveConnection = CurrentProject.Connection
    ADO_rs.Source = Dict_Source
    ADO_rs.CursorType = adOpenStatic
    ADO_rs.LockType = adLockReadOnly
    ADO_rs.Open

    X_Dict.RemoveAll

    Do While ADO_rs.EOF = False
    '添加第 0 列对应数据
        X_Dict.Add CStr(Nz(ADO_rs(ADO_rs.Fields(0).Name), "")), ""
        ADO_rs.MoveNext
    Loop
    
    Last_DictSource = Dict_Source
    
End Sub

Private Function Try_WordForm(X_Word As String, X_Pattern As String, X_Replacer As String)

    Dim X_Word_Formed As String

    If X_Pattern <> "" Then
        R_Exp.Pattern = X_Pattern
        X_Word_Formed = R_Exp.Replace(LCase(X_Word), X_Replacer)
    Else
        If X_Replacer = "Ori" Then X_Word_Formed = X_Word
        If X_Replacer = "LCase" Then X_Word_Formed = LCase(X_Word)
    End If

    If X_Dict.Exists(X_Word_Formed) Then Try_WordForm = X_Word_Formed Else Try_WordForm = ""
End Function
