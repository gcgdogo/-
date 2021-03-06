Option Explicit

Dim Value_Dictionary(2) as Object
Dim Range_Matched(2) as String
Dim Range_TobeMatched(2) as Range
Dim Range_TobeMatched_TotalRows as Long

Dim StartTimer as Double
Dim Worksheet_ForList as Worksheet , Worksheet_ForList_Finger as Long

'总程序,反复执行 Try_ValueMatch
Sub Execute_Try_ValueMatch()
    Dim I as Long
    Dim Ori_StatusBarText as Variant

    If Selection.Areas.Count<>2 Then
        Msgbox("运行需要用 Ctrl 同时选定两个区域")
        Exit Sub
    End If

    Selection.Interior.Pattern = xlNone

    Ori_StatusBarText = Application.StatusBar
    If Ori_StatusBarText = False Then Ori_StatusBarText = ""

    StartTimer = Timer()

    Set Value_Dictionary(1) = CreateObject("Scripting.Dictionary")
    Set Value_Dictionary(2) = CreateObject("Scripting.Dictionary")

    'Worksheet_ForList重置
    Set Worksheet_ForList = Nothing
    Worksheet_ForList_Finger = 0

    For I = 1 To 5000
        Range_Matched(1)=""
        Range_Matched(2)=""
        Call Try_ValueMatch

        If Range_Matched(1)="" and Range_Matched(2)="" Then
        
            Range_TobeMatched(1).Interior.Pattern = xlSolid
            Range_TobeMatched(1).Interior.Color = 255
            Range_TobeMatched(2).Interior.Pattern = xlSolid
            Range_TobeMatched(2).Interior.Color = 255

            Application.StatusBar = Ori_StatusBarText
            Msgbox("运行结束")
            Exit Sub
        End If

        If Range_TobeMatched(1) is Nothing or Range_TobeMatched(2) is Nothing Then
        
            Range_TobeMatched(1).Interior.Pattern = xlSolid
            Range_TobeMatched(1).Interior.Color = 255
            Range_TobeMatched(2).Interior.Pattern = xlSolid
            Range_TobeMatched(2).Interior.Color = 255

            Application.StatusBar = Ori_StatusBarText
            Msgbox("运行结束")
            Exit Sub
        End If
    Next
    

End Sub

'一个尝试回合,发现结果后进行标注
Private Sub Try_ValueMatch()

'Range_TobeMatched 初始设置
    If Selection.Areas.Count<>2 Then
        Msgbox("运行需要用 Ctrl 同时选定两个区域")
        Exit Sub
    End If

    Range_TobeMatched_TotalRows = 0
    Set Range_TobeMatched(1) = Get_Range_TobeMatched(Selection.Areas(1))
    Set Range_TobeMatched(2) = Get_Range_TobeMatched(Selection.Areas(2))
    'Msgbox("Range_TobeMatched: " & Union(Range_TobeMatched(1),Range_TobeMatched(2)).address)


'Value_Dictionary初始化设置
    Value_Dictionary(1).RemoveAll
    Value_Dictionary(2).RemoveAll


    Call Try_StyleA_RowsOneByOne(1)
    Call Try_StyleA_RowsOneByOne(2)


    Call Try_StyleB_FollowingLines(1)
    Call Try_StyleB_FollowingLines(2)


    Call Try_StyleC_AllCombinations(1,2)
    Call Try_StyleC_AllCombinations(2,2)
    Call Try_StyleC_AllCombinations(1,3)
    Call Try_StyleC_AllCombinations(2,3)
    Call Try_StyleC_AllCombinations(1,4)
    Call Try_StyleC_AllCombinations(2,4)
    Call Try_StyleC_AllCombinations(1,5)
    Call Try_StyleC_AllCombinations(2,5)


    If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Call Range_MarkAndList(Range(Range_Matched(1)),Range(Range_Matched(2)))
End Sub 

'第一种方式,单行遍历
Sub Try_StyleA_RowsOneByOne(Parrent_Range_Num as Integer)
    Dim Test_Area as Range , Test_Line as Range

    '若已找到匹配项 跳出
    If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Sub

    For Each Test_Area in Range_TobeMatched(Parrent_Range_Num).Areas
        For Each Test_Line in Test_Area.Rows
            'Msgbox("Test_Area: " & Test_Area.Address & "  Test_Line: " & Test_Line.Address)
            Call Try_Range_CheckAndAdd(Parrent_Range_Num , Test_Line)

            '若已找到匹配项 跳出
            If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Sub
        Next
    Next

End Sub

'第二种方式,连续行组合遍历
Sub Try_StyleB_FollowingLines(Parrent_Range_Num as Integer)
    Dim Test_Area as Range
    Dim Line_ST as Long , Line_ED as Long
    Dim CombiledLines as Range

    '若已找到匹配项 跳出
    If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Sub

    For Each Test_Area in Range_TobeMatched(Parrent_Range_Num).Areas
        For Line_ST = 1 To Test_Area.Rows.Count - 1
            For Line_ED = Line_ST + 1 To Test_Area.Rows.Count

                Set CombiledLines = Range(Test_Area.Rows(Line_ST).cells(1),Test_Area.Rows(Line_ED).cells(Test_Area.Columns.Count))

                'Msgbox("Test_Area: " & Test_Area.Address & "  CombiledLines: " & CombiledLines.Address)
                Call Try_Range_CheckAndAdd(Parrent_Range_Num , CombiledLines)

                '若已找到匹配项 跳出
                If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Sub

            Next
        Next
    Next
End Sub

'第三种方式, 不连续遍历组合
Sub Try_StyleC_AllCombinations(Parrent_Range_Num as Integer , Deep_Count as Integer)

    '若已找到匹配项 跳出
    If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Sub

    Call Try_StyleC_AllCombinations_LoopBack(Parrent_Range_Num , Nothing , Range_TobeMatched(Parrent_Range_Num) , Deep_Count , 10000)
End Sub

Function Try_StyleC_AllCombinations_LoopBack(Parrent_Range_Num as Integer , Range_Selected as Range , Range_ForSelect as Range , Deep_Count as Integer , Try_Count as Long) as Long
    Dim Finger_Area as Long, Finger_Row as Long , I as Long
    Dim Next_Area_ForSelect as Range , Next_Rows_ForSelect as Range , Next_Range_ForSelect as Range
    Dim Next_Range_Selected as Range

    '先设置好默认返回值
    Try_StyleC_AllCombinations_LoopBack = Try_Count

    If Deep_Count<0 Then Exit Function
    If Try_Count = 0 Then Exit Function


    '若已找到匹配项 跳出
    If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Function
    If Deep_Count = 0 Then 
        If not(Range_Selected is Nothing) Then Call Try_Range_CheckAndAdd(Parrent_Range_Num , Range_Selected)
        Try_StyleC_AllCombinations_LoopBack = Try_Count - 1
        Exit Function
    End If
    If Range_ForSelect is Nothing Then Exit Function
    For Finger_Area = 1 To Range_ForSelect.Areas.Count

        '若已找到匹配项 跳出
        If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Function

        '从Area+1开始循环, 设定 Next_Area_ForSelect
        Set Next_Area_ForSelect = Nothing
        For I = Finger_Area + 1 To Range_ForSelect.Areas.Count
            If Next_Area_ForSelect is Nothing Then
                Set Next_Area_ForSelect = Range_ForSelect.Areas(I)
            Else
                Set Next_Area_ForSelect = Union(Next_Area_ForSelect , Range_ForSelect.Areas(I))
            End If
        Next

        For Finger_Row = 1 To Range_ForSelect.Areas(Finger_Area).Rows.Count

        '若已找到匹配项 跳出
        If Range_Matched(1)<>"" and Range_Matched(2)<>"" Then Exit Function

            '从Row+1开始循环, 设定 Next_Range_ForSelect
            Set Next_Rows_ForSelect = Nothing
            For I = Finger_Row+1 To Range_ForSelect.Areas(Finger_Area).Rows.Count
                If Next_Rows_ForSelect is Nothing Then
                    Set Next_Rows_ForSelect = Range_ForSelect.Areas(Finger_Area).Rows(I)
                Else
                    Set Next_Rows_ForSelect = Union(Next_Rows_ForSelect , Range_ForSelect.Areas(Finger_Area).Rows(I))
                End If
            Next

            '组合得出 Next_Range_ForSelect
            Set Next_Range_ForSelect = Next_Area_ForSelect
            If Next_Range_ForSelect is Nothing Then
                Set Next_Range_ForSelect = Next_Rows_ForSelect
            Else
                If Not(Next_Rows_ForSelect is Nothing) Then Set Next_Range_ForSelect = Union(Next_Range_ForSelect , Next_Rows_ForSelect)
            End if

            '组合得出 Next_Range_Selected
            Set Next_Range_Selected = Range_ForSelect.Areas(Finger_Area).Rows(Finger_Row)
            If Not(Range_Selected is Nothing) Then Set Next_Range_Selected = Union(Next_Range_Selected , Range_Selected)

            '进行递归调用
            Try_Count = Try_StyleC_AllCombinations_LoopBack(Parrent_Range_Num , Next_Range_Selected , Next_Range_ForSelect , Deep_Count - 1 , Try_Count)

        Next
    Next

    '返回 Try_Count
    Try_StyleC_AllCombinations_LoopBack = Try_Count

End Function



'监测该项是否能与对面匹配,无匹配则增加至 Dictionary
Sub Try_Range_CheckAndAdd(Parrent_Range_Num as Integer , X_Range as Range)
    Static LastTimer_DoEvent as Double , Check_Count as Long , Check_Count_inTime as Double
    Dim X_Range_Sum as Long

    Check_Count = Check_Count + 1
    Check_Count_inTime = Check_Count_inTime + 1

    If LastTimer_DoEvent < StartTimer Then LastTimer_DoEvent = StartTimer

    '定期 DoEvents
    If Timer() - LastTimer_DoEvent > 1 Then
        DoEvents
        Application.StatusBar = "待匹配条目 : " & Range_TobeMatched_TotalRows & " | 已尝试组合 : " & Check_Count & " ( " & Int(Check_Count_inTime / (Timer() - LastTimer_DoEvent)) & "/s )"
        LastTimer_DoEvent=Timer()
        Check_Count_inTime = 0
    End If

    X_Range_Sum = Get_Range_ValueSum(X_Range)

    If Value_Dictionary(3 - Parrent_Range_Num).Exists(X_Range_Sum) Then
        Range_Matched(Parrent_Range_Num)=X_Range.Address
        Range_Matched(3 - Parrent_Range_Num)=Value_Dictionary(3 - Parrent_Range_Num)(X_Range_Sum)
        'Msgbox("Range_Matched_1: " & Range_Matched(1) & vbCrLf & "Range_Matched_2: " & Range_Matched(2) )
    Else
        If Value_Dictionary(Parrent_Range_Num).Exists(X_Range_Sum)=False Then
            Value_Dictionary(Parrent_Range_Num).Add X_Range_Sum , X_Range.Address
        End If
    End If
End Sub


'获得需要进行匹配的Range
Function Get_Range_TobeMatched(X_Range as Range) as Range
    Dim Got_Range as Range, Test_Area as Range , Test_Line as Range
    For Each Test_Area in X_Range.Areas
        For Each Test_Line in Test_Area.Rows
            If Test_Line.Interior.Pattern = xlNone and Application.WorksheetFunction.IsNumber(Test_Line.Cells(Test_Line.Cells.Count)) Then

                Range_TobeMatched_TotalRows = Range_TobeMatched_TotalRows + 1

                If Got_Range is Nothing Then
                    Set Got_Range = Test_Line
                Else
                    Set Got_Range = Union(Got_Range , Test_Line)
                End If
            End If
        Next
    Next

    Set Get_Range_TobeMatched = Got_Range
End Function

'为避免浮点误差, Get_Range_ValueSum 计算结果为 Int(Sum(Range)*100)
Function Get_Range_ValueSum(X_Range as Range) as Long
    Dim X_Sum as Long
    Dim Test_Area as Range , Test_Line as Range

    X_Sum=0
    For Each Test_Area in X_Range.Areas
        For Each Test_Line in Test_Area.Rows
            X_Sum = X_Sum + Int(0.5 + Test_Line.Cells(Test_Line.Cells.Count) * 100)
        Next
    Next

    Get_Range_ValueSum=X_Sum
End Function


'匹配项标注并单独列出
Sub Range_MarkAndList(Range_A as Range , Range_B as Range)

    Dim Range_Color as Long
    Dim Active_Sheet_Name as String

    Dim Active_Finger_A as Long , Active_Finger_B as Long
    Dim Test_Area as Range , Test_Line as Range , Cell_Num as Integer
    Dim Range_B_Start_Column as Long
    If Worksheet_ForList is Nothing Then
        Active_Sheet_Name = Activesheet.Name
        Worksheets.Add
        Set Worksheet_ForList = Worksheets(Activesheet.Name)
        Worksheets(Active_Sheet_Name).Activate
        Worksheet_ForList_Finger=1
    End If

    '根据结果区域连续性确认颜色
    '默认涂色为浅绿色
    Range_Color = 10747835
    '连续多行为浅蓝色
    If Range_A.Rows.Count + Range_B.Rows.Count > 2 Then Range_Color = 16770749
    '多个Area为浅黄色
    If Range_A.Areas.Count + Range_B.Areas.Count > 2 Then Range_Color = 11010047
    
    Range_A.Interior.Pattern = xlSolid
    Range_A.Interior.Color = Range_Color
    Range_B.Interior.Pattern = xlSolid
    Range_B.Interior.Color = Range_Color


    Active_Finger_A = Worksheet_ForList_Finger

    For Each Test_Area in Range_A.Areas
        For Each Test_Line in Test_Area.Rows
            For Cell_Num = 1 to Test_Line.Cells.Count
                Worksheet_ForList.Rows(Active_Finger_A).Cells(Cell_Num) = Test_Line.Cells(Cell_Num)
                Worksheet_ForList.Rows(Active_Finger_A).Cells(Cell_Num).Interior.Pattern = xlSolid
                Worksheet_ForList.Rows(Active_Finger_A).Cells(Cell_Num).Interior.Color = Range_Color
            Next
            Active_Finger_A=Active_Finger_A+1
        Next
    Next

    Active_Finger_B = Worksheet_ForList_Finger
    Range_B_Start_Column = Range_A.Columns.Count + 1

    For Each Test_Area in Range_B.Areas
        For Each Test_Line in Test_Area.Rows
            For Cell_Num = 1 to Test_Line.Cells.Count
                Worksheet_ForList.Rows(Active_Finger_B).Cells(Cell_Num + Range_B_Start_Column) = Test_Line.Cells(Cell_Num)
                Worksheet_ForList.Rows(Active_Finger_B).Cells(Cell_Num + Range_B_Start_Column).Interior.Pattern = xlSolid
                Worksheet_ForList.Rows(Active_Finger_B).Cells(Cell_Num + Range_B_Start_Column).Interior.Color = Range_Color
            Next
            Active_Finger_B = Active_Finger_B+1
        Next
    Next

    Worksheet_ForList_Finger = IIF(Active_Finger_A > Active_Finger_B , Active_Finger_A , Active_Finger_B) + 1
    'Msgbox("Range_A: " & Range_A.Address & vbCrLf & "Range_B: " & Range_B.Address )
End Sub