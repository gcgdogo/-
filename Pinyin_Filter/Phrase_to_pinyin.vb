Option Explicit
'可以使用UseDict启用是否切换为使用 X_Dict 保存已计算拼音

Function Phrase_to_pinyin(P as String , Optional UseDict as Boolean = False) as String
    Dim I as Integer , J as Integer 
    Dim Pinyin as String,Pinyin_Temp as String
    Dim RegExp_Replace as new RegExp
    Dim RegExp_Match as new RegExp
    Dim R_Matches as MatchCollection

    Static X_Dict as Object
    Dim X_Dict_Found as String

    ' 如果 UseDict 判定是否存在
    If UseDict=True Then
        If X_Dict is Nothing Then Set X_Dict = CreateObject("Scripting.Dictionary") 
        X_Dict_Found = X_Dict(P)
        If X_Dict_Found<>"" Then
            Phrase_to_pinyin = X_Dict_Found
            Exit Function
        End If
    End If

    Pinyin=","
    
    RegExp_Replace.IgnoreCase = True
    RegExp_Replace.Global = True
    RegExp_Replace.Pattern=","

    RegExp_Match.IgnoreCase=True
    RegExp_Match.Global=True
    RegExp_Match.Pattern="[^,]*,"

    For I=1 To Len(P)
        set R_Matches = RegExp_Match.Execute(Character_to_Pinyin(Mid(P,I,1)))
        Pinyin_Temp=""
        'MatchCollection 的集合从0开始到n-1
        For J=0 To R_Matches.Count-1
            Pinyin_Temp = Pinyin_Temp & RegExp_Replace.Replace(Pinyin,R_Matches(J).Value)
        Next
        Pinyin=Pinyin_Temp
        'Msgbox("I=" & I & " : " & Pinyin)
    Next

    '用来检测 X_Dict 保存情况
    ' Pinyin = Pinyin & X_Dict.Count

    ' 如果 UseDict 储存结果
    If UseDict=True Then X_Dict(P) = Pinyin

    Phrase_to_pinyin=Pinyin
End Function