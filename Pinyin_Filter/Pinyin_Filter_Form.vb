Option Explicit

Dim X_Phrase() As String, X_Text() As String, X_Count As Long
Dim FilterRange As Range, FilterNum As Long

'_________________________________________________Initalize
'直接再Initalize事件中关闭窗口会报错,所以在激活时检验是否数据已经初始化
Private Sub UserForm_Activate()
    If X_Count = 0 Then
        MsgBox ("筛选数据未初始化")
        Unload Me
    End If
        
End Sub

Private Sub UserForm_Initialize()
    Dim StartTime as Double
    Dim X_Cell as Range
    Dim I as Long
    StartTime=Timer()
    X_Count=0
    If ActiveSheet.AutoFilterMode = False Then
        MsgBox ("需要启用筛选")
        Exit Sub
    End If
    Set FilterRange = ActiveSheet.AutoFilter.Range
    FilterNum = Selection.Cells(1).Column - FilterRange.Cells(1).Column + 1
    If FilterRange.Columns(FilterNum).Cells.Count > 50000 Then
        MsgBox ("筛选范围过大 [>50000]")
        Exit Sub
    End If

    For Each X_Cell In FilterRange.Columns(FilterNum).cells
        For I=1 To X_Count
            If X_Phrase(I)=X_Cell.Value Then Exit For         
        Next
        If I>X_Count Then
            Redim Preserve X_Phrase(I) As String, X_Text(I) As String
            X_Phrase(I) = X_Cell.Value
            X_Text(I) = Left(X_Cell.Value & "　　　　" ,IIF(Len(X_Cell.Value)>4,Len(X_Cell.Value),4)) & ": " & Phrase_to_pinyin(X_Cell.Value)
            X_Count=I
        End If
    Next

    Call Show_TimeElapsed(Timer()-StartTime , "  已读取拼音数量 = " & X_Count)

End Sub

'_________________________________________________TextBox_Pinyin

Private Sub TextBox_Pinyin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'Enter = 全部导入
    If KeyCode = 13 and Shift = 0 Then Call Button_ChooseAll_Click
    'Ctrl+Enter = 执行
    If KeyCode = 13 and Shift = 2 Then Call Button_Execute_Click

End Sub


Private Sub TextBox_Pinyin_Change()
    Dim StartTime as Double
    Dim I as Long
    Dim RegExp_Chk as new RegExp
    If Me.TextBox_Pinyin.Value = "" Then Exit Sub
    StartTime=Timer()

    RegExp_Chk.Global = True
    RegExp_Chk.IgnoreCase = True
    '匹配除最后一个字符外的全部字符, 使用零宽度正预测先行断言
    RegExp_Chk.Pattern = "(.)(?=.)"

    Me.ListBox_Optional.Clear

    '先按大写首字母测试 , 不能有多余大写字母
    RegExp_Chk.Pattern = "[^a-zA-Z]" & RegExp_Chk.Replace(Ucase(TextBox_Pinyin.Value) , "$1[^,:A-Z]*") & "[^A-Z]*,"
    RegExp_Chk.IgnoreCase = False
    For I = 1 To X_Count
        If RegExp_Chk.Test(X_Text(I))=True Then ListBox_Optional.AddItem(X_Text(I))
    Next

    '大写首字母没通过就不管大写重新测试
    If ListBox_Optional.ListCount=0 Then
        RegExp_Chk.Pattern = "(.)(?=.)"
        RegExp_Chk.Pattern = RegExp_Chk.Replace(TextBox_Pinyin.Value , "$1[^,:]*")
        RegExp_Chk.IgnoreCase = True
        For I = 1 To X_Count
            If RegExp_Chk.Test(X_Text(I))=True Then ListBox_Optional.AddItem(X_Text(I))
        Next
    End If

    Call Show_TimeElapsed(Timer()-StartTime , " Exp = " & RegExp_Chk.Pattern)

End Sub

'_________________________________________________ListBox
Private Sub ListBox_Confirmed_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ListBox_Confirmed.RemoveItem (Me.ListBox_Confirmed.ListIndex)
    '让输入框保持焦点
    Me.TextBox_Pinyin.SetFocus
End Sub

Private Sub ListBox_Optional_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ListBox_Confirmed.AddItem (Me.ListBox_Optional.Value)
    '清空 并让输入框保持焦点
    Me.TextBox_Pinyin.Text = ""
    Me.TextBox_Pinyin.SetFocus
End Sub

'_________________________________________________Button

Private Sub Button_ChooseAll_Click()
    Dim I as Long
    For I = 1 To ListBox_Optional.ListCount
        'ListBox.List 列表从零开始
        Me.ListBox_Confirmed.AddItem (Me.ListBox_Optional.List(I-1))
    Next
End Sub

Private Sub Button_Execute_Click()
    Dim I as Long
    Dim RegExp_Match as new RegExp
    Dim X_Criteria() as String

    RegExp_Match.Global=False
    RegExp_Match.IgnoreCase=True
    RegExp_Match.Pattern="^[^　:]*"

    Redim X_Criteria(Me.ListBox_Confirmed.ListCount-1) as String

    For I=0 To Me.ListBox_Confirmed.ListCount-1
        X_Criteria(I) = "=" & RegExp_Match.Execute(Me.ListBox_Confirmed.List(I)).Item(0).Value
    Next

    'Msgbox(I)

    FilterRange.AutoFilter field:=FilterNum , Criteria1:= X_Criteria , Operator:=xlFilterValues

    Unload Me

End Sub

'_________________________________________________TimeElapsed

Private Sub Show_TimeElapsed(T As Double , Optional Addon_String as String = "")
    T = Int(T * 1000)
    Me.Label_TimeElapsed.Caption = "[Time Elapsed] = " & T & "ms" & Addon_String
End Sub
