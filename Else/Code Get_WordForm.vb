Option Explicit

Dim X_Dict as new Dictionary
Dim TimerFlag as Double
Dim R_Exp as new RegExp

Public Function Get_WordForm(X_Word as String , Dict_Source as String)
    
    '如果距离上次执行超过规定时间就重置词典
    If Timer()-TimerFlag > 10 Then Call Get_WordForm_LoadDict(Dict_Source)

    Get_WordForm = Try_WordForm(X_Word , "" , "Ori")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "" , "LCase")

    '加s系列
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "s$" , "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "es$" , "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ies$" , "y")

    '加ing系列
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ing$" , "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ing$" , "e")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ying$" , "ie")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "(.)\1ing$" , "$1")

    '加ed系列
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ed$" , "")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ed$" , "e")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "ied$" , "y")
    If Get_WordForm = "" Then Get_WordForm = Try_WordForm(X_Word , "(.)\1ed$" , "$1")

    TimerFlag = Timer()

End Function


Private Sub Get_WordForm_LoadDict(Dict_Source as String)
    Dim ADO_rs as new ADODB.Recordset
    ADO_rs.ActiveConnection = CurrentProject.Connection
    ADO_rs.Source = Dict_Source
    ADO_rs.CursorType = adOpenStatic
    ADO_rs.LockType = adLockReadOnly
    ADO_rs.Open

    X_Dict.Removeall

    Do While ADO_rs.EOF = False
    '添加第 0 列对应数据
        X_Dict.Add CStr(nz(ADO_rs(ADO_rs.Fields(0).Name),"")) , ""
        ADO_rs.MoveNext
    Loop

End Sub

Private Function Try_WordForm(X_Word as String , X_Pattern as String , X_Replacer as String)

    Dim X_Word_Formed as String

    If X_Pattern<>"" Then
        R_Exp.Pattern = X_Pattern
        X_Word_Formed = R_Exp.Replace(LCase(X_Word) , X_Replacer)
    Else
        If X_Replacer = "Ori" Then X_Word_Formed=X_Word
        If X_Replacer = "LCase" Then X_Word_Formed=LCase(X_Word)
    End If

    If X_Dict.Exists(X_Word_Formed) Then Try_WordForm = X_Word_Formed Else Try_WordForm = ""
End Function