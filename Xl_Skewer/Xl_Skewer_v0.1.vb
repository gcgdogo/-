Option Explicit

Dim Global_SkewerWorkbook as Workbook
Dim Global_SkewerHeader as Range

'存储相对路径
Dim Waker_Filenames(1000) as String , Waker_Filenames_Count as Integer , Waker_Filenames_Finger as Integer

'设定Header
Sub Skewer_HeaderSet(optional Cal As XlCalculation)
    Dim Range_SkewerHeader as Range
    Dim I as Integer , I_Max as Integer
    Dim Range_Check as Boolean

    Set Range_SkewerHeader = Selection.Cells(1)

    '暂时只支持首列起点
    Range_Check = ( Range_SkewerHeader.Column = 1 )
    '检测当前行是否为空
    I_Max = ActiveSheet.UsedRange.Columns.Count

    For I = 0 to I_Max
        if Range_SkewerHeader.offset(0,I).Text <> "" then Range_Check = False
    Next

    If Range_Check = False then
        Msgbox("Range_Check 失败\n自动选择选区第一格\n位置要为首列\n需要整行为空")
        Exit Sub
    End if

    '如果校验成功，记录位置
    set Range_SkewerHeader = range(Range_SkewerHeader , Range_SkewerHeader.offset(0,3))

    '如果没有 Names("SkewerHeader") 就新建一个
    For I = 1 To ActiveWorkbook.Names.Count
        if ActiveWorkbook.Names(I).Name = "SkewerHeader" then exit for
    Next
    If I > ActiveWorkbook.Names.Count then ActiveWorkbook.Names.Add Name:="SkewerHeader", RefersTo:="=""none"""

    ActiveWorkbook.Names("SkewerHeader").RefersTo = "='" & ActiveSheet.Name & "'!" & Range_SkewerHeader.Address
    
    '清除已有 Range_SkewerHeader 标题的加粗格式
    I_Max = ActiveSheet.UsedRange.Rows.Count

    For I = 0 to I_Max
        With Range("A1:D1").offset(I,0)
            if .Cells(1).Text = "FileName" and _
                .Cells(2).Text = "Label" and _
                .Cells(3).Text = "Sheet" and _
                .Cells(4).Text = "Range" _
            then .Font.Bold = False
        End With
    Next

    '格式化 Range_SkewerHeader
    With Range_SkewerHeader.EntireRow
        .Clear
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlThin
    End With

    Range_SkewerHeader.Cells(1).Value = "FileName"
    Range_SkewerHeader.Cells(2).Value = "Label"
    Range_SkewerHeader.Cells(3).Value = "Sheet"
    Range_SkewerHeader.Cells(4).Value = "Range"

    if Columns("A:A").ColumnWidth < 36 then Columns("A:A").ColumnWidth = 36
    if Columns("B:B").ColumnWidth < 18 then Columns("B:B").ColumnWidth = 18
    if Columns("C:C").ColumnWidth < 18 then Columns("C:C").ColumnWidth = 18
    if Columns("D:D").ColumnWidth < 18 then Columns("D:D").ColumnWidth = 18
End Sub


'根据已打开表格，定位Skewer位置
Function Find_SkewerWorkbook() as Boolean
    Dim X_Workbook as Workbook , X_Names as Name
    Dim SkewerWorkbook_Found as Boolean
    SkewerWorkbook_Found = False
    for each X_Workbook in application.workbooks
        for each X_Names in X_Workbook.Names
            if X_Names.Name = "SkewerHeader" then
                if SkewerWorkbook_Found then
                    Msgbox("无法定位：发现两个文件存在SkewerHeader：[ " & Global_SkewerWorkbook.Name & " ] [ " & X_Workbook.Name &" ]")
                    Find_SkewerWorkbook = False
                    Exit Function
                else
                    Set Global_SkewerWorkbook = X_Workbook
                    SkewerWorkbook_Found = True
                end if
            end if
        next
    next

    'SkewerHeader 必须已经保存，即为有路径
    if SkewerWorkbook_Found then
        if Global_SkewerWorkbook.Path = "" then
            Msgbox("无法定位：SkewerHeader所在路径获取失败")
            Find_SkewerWorkbook = False
            Exit Function
        end If
    end If

    if not SkewerWorkbook_Found then
        Msgbox("无法定位：未发现文件存在SkewerHeader")
        Find_SkewerWorkbook = False
        Exit Function
    end if

    Set Global_SkewerHeader = Global_SkewerWorkbook.Names("SkewerHeader").RefersToRange
    Find_SkewerWorkbook = True
End Function

'获取表长度
Function SkewerFinger_MaxRow()
    Dim SkewerFinger_Row as Long
    Static SkewerFinger_MaxRow_Stock as Long , SkewerFinger_MaxRow_CheckString as String
    Dim X_Cell as Range , X_CheckString as String
    Dim X_FileName as String , X_FileName_Check as Boolean
    '先看看末尾值是否已经调整
    X_CheckString = ""
    for each X_Cell in Union(Global_SkewerHeader.offset(SkewerFinger_MaxRow_Stock,0),Global_SkewerHeader.offset(SkewerFinger_MaxRow_Stock + 1,0))
        X_CheckString = X_CheckString & ":" & X_Cell.Text
    next
    '校验通过，那我不改了
    if X_CheckString = SkewerFinger_MaxRow_CheckString then
        SkewerFinger_MaxRow = SkewerFinger_MaxRow_Stock
        Exit Function
    End if
    for SkewerFinger_Row = 1 to 500000  '先按照50万行最大值循环
        X_FileName_Check = False 
        X_FileName = Global_SkewerHeader.offset(SkewerFinger_Row,0).Cells(1).Text
        if right(X_FileName,4) = ".xls" then X_FileName_Check = True
        if right(X_FileName,5) = ".xlsx" then X_FileName_Check = True
        if right(X_FileName,5) = ".xlsm" then X_FileName_Check = True  '暂时就先支持这几种格式吧
        if X_FileName_Check = False then
            SkewerFinger_MaxRow = SkewerFinger_Row - 1
            exit Function
        end If
    next
    msgbox("Function SkewerFinger_MaxRow() 获取失败")

End Function

'读取待处理文件名
Sub Skewer_LoadFilename(optional Cal As XlCalculation)
    Dim X_Workbook as Workbook , X_Workbook_PathStock as String , X_Workbook_PathSame as Boolean
    Dim SkewerPath as String
    Dim SkewerFinger_Row as Long , X_Workbook_Found as Boolean , X_Workbook_CutName as String

    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    '检查现有文件路径，是否是子文件夹
    SkewerPath = Global_SkewerWorkbook.Path
    X_Workbook_PathStock = ""
    X_Workbook_PathSame = True
    for each X_Workbook in application.workbooks
        if X_Workbook.Name <> Global_SkewerWorkbook.Name then
            '是否是子文件夹
            if left(X_Workbook.Path,len(SkewerPath)) <> SkewerPath then
                msgbox("路径校验失败：有文件不在SkewerWorkbook的子目录下：" & X_Workbook.FullName)
                exit Sub
            end if
            '是否为同一文件夹
            if X_Workbook_PathStock ="" then X_Workbook_PathStock = X_Workbook.Path
            if X_Workbook_PathStock <> X_Workbook.Path then X_Workbook_PathSame = False
        end if
    next
    '检查现有文件是否为同一文件夹
    if X_Workbook_PathSame = False then msgbox("提示：文件不在同文件夹，建议检查文件")

    '遍历，文件
    for each X_Workbook in application.workbooks
        if X_Workbook.Name <> Global_SkewerWorkbook.Name then
            '检查文件名是否已存在
            X_Workbook_Found = False
            X_Workbook_CutName = mid(X_Workbook.FullName , len(Global_SkewerWorkbook.Path) + 1 , 9999)
            for SkewerFinger_Row = 1 to SkewerFinger_MaxRow()
                if  X_Workbook_CutName = Global_SkewerHeader.offset(SkewerFinger_Row,0).Cells(1).Text then
                    X_Workbook_Found = True
                    exit for
                end if
            next
            '没找到就写入文件名
            if X_Workbook_Found = False then
                Global_SkewerHeader.offset(SkewerFinger_MaxRow() + 1,0).Cells(1).Value = X_Workbook_CutName
            end if
        '关闭文件
        X_Workbook.Close
        end if
    next
End Sub

'开始遍历目标文件
Sub Skewer_FileWalker_Start(optional Cal As XlCalculation)
    Dim SkewerFinger_Row as Long , SkewerFinger_Filename as String
    Dim I as Integer

    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    '当前文件必须是 SkewerWorkbook
    If ActiveWorkbook.Name <> Global_SkewerWorkbook.Name then
        msgbox("起点错误：当前文件必须是 Skewer")
        exit Sub
    end If

    Waker_Filenames_Count = 0 '重置 Waker_Filenames
    Waker_Filenames_Finger = 0

    '遍历文件名
    for SkewerFinger_Row = 1 to SkewerFinger_MaxRow()
        SkewerFinger_Filename = Global_SkewerHeader.offset(SkewerFinger_Row,0).Cells(1).Text
        '遍历核对现有 Waker_Filenames
        for I = 1 to Waker_Filenames_Count
            if  SkewerFinger_Filename = Waker_Filenames(I) then exit for
        next
        '没有则新增
        if I > Waker_Filenames_Count then
            Waker_Filenames_Count = Waker_Filenames_Count + 1
            Waker_Filenames(Waker_Filenames_Count) = SkewerFinger_Filename
        end if
    next

    Call Skewer_FileWalker_Next '调用打开文件

End Sub

'打开下一个文件
'反馈是否成功打开
Function Skewer_FileWalker_Next() as Boolean           
    Dim X_FullName as String
    Dim X_Workbook as Workbook

    Skewer_FileWalker_Next = False '初始值
    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Function

    'Finger自增，检验是否已经结尾
    Waker_Filenames_Finger = Waker_Filenames_Finger + 1
    if Waker_Filenames_Finger > Waker_Filenames_Count then
        msgbox("文件遍历结束")
        exit Function
    end if

    X_FullName = Global_SkewerWorkbook.Path + Waker_Filenames(Waker_Filenames_Finger)

    '检测是否有同名文件
    for each X_Workbook in application.workbooks
        if right(X_FullName,len(X_Workbook.Name) + 1) = "\" & X_Workbook.Name then
            msgbox("检测到同名文件，需要将其关闭")
            X_Workbook.Close
        end if
    next
    '如果关闭失败，终止
    for each X_Workbook in application.workbooks
        if right(X_FullName,len(X_Workbook.Name) + 1) = "\" & X_Workbook.Name then
            msgbox("同名文件关闭失败，程序终止")
            Skewer_FileWalker_Next = False
            exit function
        end if
    next

    '以只读方式打开文件
    application.workbooks.open Filename:=X_FullName , ReadOnly:=True

    if activeworkbook.FullName <> X_FullName then
        msgbox("文件名不一致，目标文件打开失败，程序终止")
        Skewer_FileWalker_Next = False
        exit function
    end if

    '调用格式化 Skewer_Range_Marker
    call Skewer_Range_Marker

    Skewer_FileWalker_Next = True

End Function

'设定Sheet,Range
Sub Skewer_Range_Set(optional Cal As XlCalculation)
    Dim X_Worksheet as Worksheet
    Dim SkewerFinger_Row as Long

    Dim X_CutName as String, X_Sheet as String, X_Range as String

    Dim X_Range_RowED as Long , X_Range_ColED as Long
    Dim X_Cell as Range

    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    X_CutName = mid(activeworkbook.FullName , len(Global_SkewerWorkbook.Path) + 1 , 9999)


    For each X_Worksheet in ActiveWindow.SelectedSheets
        X_Sheet = X_Worksheet.Name
        '检测现有Range
        X_Range = ""
        for SkewerFinger_Row = 1 to SkewerFinger_MaxRow()
            With Global_SkewerHeader.offset(SkewerFinger_Row,0)
                if .Cells(1).Text = X_CutName and .Cells(3).Text = X_Sheet then
                    X_Range =  .Cells(4).Text
                    exit for
                end if
            End With
        next

        '如果已有Range设定 ActiveWindow.SelectedSheets 只有一个，且Selection所选择区域不大不小，就把选区设为Range
        '选区大小改为之后校验
        If X_Range <> "" and ActiveWindow.SelectedSheets.Count = 1 then
            X_Range = Selection.Address
        else
            '默认采用ActiveSheet.UsedRange遍历删掉末尾无效行列的方式
            X_Range_RowED = 1
            X_Range_ColED = 1
            for each X_Cell in X_Worksheet.UsedRange.Cells
                if X_Cell.Text <> "" or X_Cell.MergeCells then
                    if X_Cell.Row > X_Range_RowED then X_Range_RowED = X_Cell.Row
                    if X_Cell.Column > X_Range_ColED then X_Range_ColED = X_Cell.Column
                end if
            next
            X_Range = "$A$1:" & X_Worksheet.cells( X_Range_RowED , X_Range_ColED ).address
        end If

        With X_Worksheet.Range(X_Range)
            if .cells.count < 5 then
                msgbox("选区过小保护(<5)")
                X_Range = ""
            end if
            if .Columns.Count > 100 then
                msgbox("选区过宽保护(>100)")
                X_Range = ""
            end If
            if .Rows.Count > 10000 then
                msgbox("选区过长保护(>10000)")
                X_Range = ""
            end If
        End With
        Call Write_Skewer_RangeData(activeworkbook.FullName , X_Worksheet.Name , X_Range)
    Next

    '调用Skewer_Range_Marker
    Call Skewer_Range_Marker

End Sub

'写入一行数据 , X_Range如果是空的就光是删除数据
Sub Write_Skewer_RangeData(X_FullName as String, X_Sheet as String, X_Range as String)
    Dim X_CutName as String , X_Label as String
    Dim SkewerFinger_Row as Long , SkewerFinger_Range as Range
    Dim SkewerFinger_SelectedRow as Long
    Dim Range_RowsToDelete as Range , Range_RowsToDelete_IsSet as Boolean

    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    X_CutName = mid(X_FullName , len(Global_SkewerWorkbook.Path) + 1 , 9999)
    X_Label = ""
    
    '遍历删除所有对应 X_Sheet 的数据
    Range_RowsToDelete_IsSet = False
    SkewerFinger_SelectedRow =0

    for SkewerFinger_Row = SkewerFinger_MaxRow to 1 step -1
        set SkewerFinger_Range = Global_SkewerHeader.offset(SkewerFinger_Row,0)

        if SkewerFinger_Range.Cells(1).Text = X_CutName and SkewerFinger_SelectedRow = 0 then '如果没有设置，就找文件名之后的一行
            SkewerFinger_SelectedRow = SkewerFinger_Row + 1
        end If

        if SkewerFinger_Range.Cells(1).Text = X_CutName and SkewerFinger_Range.Cells(3).Text = X_Sheet then '判定一致

            SkewerFinger_SelectedRow = SkewerFinger_Row
            X_Label = SkewerFinger_Range.Cells(2).Text '记录 X_Label

            if Range_RowsToDelete_IsSet then '记录需要删除的行
                set Range_RowsToDelete = Union(Range_RowsToDelete , SkewerFinger_Range.EntireRow)
            else
                set Range_RowsToDelete = SkewerFinger_Range.EntireRow
                Range_RowsToDelete_IsSet = true
            end if
        end if

        if SkewerFinger_Range.Cells(1).Text = X_CutName and SkewerFinger_Range.Cells(3).Text = "" and X_Label = "" then
            X_Label = SkewerFinger_Range.Cells(2).Text  '记录 X_Label
        end if
    next

    if Range_RowsToDelete_IsSet then Range_RowsToDelete.Delete  '删除需要删除的行

    'X_Range如果是空的就光是删除数据
    if X_Range <> "" then
        '新建一行
        Global_SkewerHeader.offset(SkewerFinger_SelectedRow,0).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        With Global_SkewerHeader.offset(SkewerFinger_SelectedRow,0)
            .cells(1) = X_CutName
            .cells(3) = X_Sheet
            .cells(4) = X_Range
        End With
    End If
End Sub

'删除区域选定
Sub Skewer_Range_UnSelect(optional Cal As XlCalculation)
    Dim X_Worksheet as Worksheet
    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    For each X_Worksheet in ActiveWindow.SelectedSheets
        '遍历删除该表对应的所有行
        Call Write_Skewer_RangeData(activeworkbook.FullName , X_Worksheet.Name , "") '设为空就删除了
    Next

    '调用Skewer_Range_Marker
    Call Skewer_Range_Marker

End Sub



'标注已选定的区域
Sub Skewer_Range_Marker()
    Dim SkewerFinger_Row as Long , X_CutName as String
    Dim I as Integer , J as Integer
    Dim TabColor(5) as Integer , TabColor_Str as String
    Dim ActiveSheet_Name as String
    Dim Edit_Sheet_Name as String , Edit_Range_Address as String
    
    Dim Edit_Range_LastCell as Range , Split_Range_Height as Double
    Dim Remark_All_Sheets as Boolean

    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub
    
    '验证当前文件，是否只读，是否在列表内，是否为Skewer
    if activeworkbook.ReadOnly = False then 
        msgbox("当前文件不是只读文件，无法渲染")
        exit Sub
    end if

    if activeworkbook.name = Global_SkewerWorkbook.name then 
        msgbox("当前文件是 SkewerWorkbook ，无法渲染")
        exit Sub
    end if

    X_CutName = mid(X_FullName , len(Global_SkewerWorkbook.Path) + 1 , 9999)
    for SkewerFinger_Row = 1 to SkewerFinger_MaxRow()
        if Global_SkewerHeader.offset(SkewerFinger_Row,0).Cells(1).Text = X_CutName then
            exit for
        end if
    next
    if SkewerFinger_Row > SkewerFinger_MaxRow() then
        msgbox("当前文件没在 Skewer 范围内记录，无法渲染")
        exit Sub
    end if

    '在Names里记录下各个Sheet原始标签颜色
    for I = 1 to activeworkbook.names.count  '先判断一下是不是已经记录了颜色
        if activeworkbook.names(I).name = "Skewer_SheetColor" then exit for
    next

    Remark_All_Sheets = False

    if I > activeworkbook.names.count then '如果没有记录颜色就记录一下

        Remark_All_Sheets = True '如果没有颜色那就全重新标记一遍！

        TabColor_Str = ""
        for I = 1 to Worksheets.Count
            if Worksheets(I).Tab.Color = False then
                TabColor(1) = 255
                TabColor(2) = 255
                TabColor(3) = 255
                TabColor(4) = 255
            else
                TabColor(1) = int((int( Worksheets(I).Tab.Color / 256 / 256 / 256))/2 + 128)
                TabColor(2) = int((int( Worksheets(I).Tab.Color / 256 / 256) mod 256)/2 + 128)
                TabColor(3) = int((int( Worksheets(I).Tab.Color / 256 ) mod 256)/2 + 128)
                TabColor(4) = int((int( Worksheets(I).Tab.Color ) mod 256)/2 + 128)
            end if

            TabColor_Str = TabColor_Str & "[" & Worksheets(I).name
            TabColor_Str = TabColor_Str & ":" & right("0000" & str(TabColor(1)),3)
            TabColor_Str = TabColor_Str & ":" & right("0000" & str(TabColor(2)),3)
            TabColor_Str = TabColor_Str & ":" & right("0000" & str(TabColor(3)),3)
            TabColor_Str = TabColor_Str & ":" & right("0000" & str(TabColor(4)),3)
            TabColor_Str = TabColor_Str & "] "
        next
        ActiveWorkbook.Names.Add Name:="Skewer_SheetColor", RefersTo:="=""" & TabColor_Str & """"
    end if

    '保存Actviesheet 便于执行完毕后恢复
    ActiveSheet_Name = ActiveSheet.Name

  
    for I = 1 to Worksheets.Count
        if Remark_All_Sheets = True or Worksheets(I).Tab.Color > 0 then  '如果不是 Remark_All_Sheets 就不重置已经是黑色的标签的颜色
            for J = 1 to len(ActiveWorkbook.Names("Skewer_SheetColor").Value)
                '把标签全都重置成原颜色的浅色版本
                TabColor_Str = mid(ActiveWorkbook.Names("Skewer_SheetColor").Value , J , len(Worksheets(I).name) + 2 + 16)
                if left( TabColor_Str , len(Worksheets(I).name) + 2) = "[" & Worksheets(I).name & ":" then
                    TabColor_Str = mid(TabColor_Str , len(Worksheets(I).name) + 3 , 15)
                    TabColor(1) = Val(mid(TabColor_Str , 1 , 3))
                    TabColor(2) = Val(mid(TabColor_Str , 5 , 3))
                    TabColor(3) = Val(mid(TabColor_Str , 9 , 3))
                    TabColor(4) = Val(mid(TabColor_Str , 13 , 3))
                    Worksheets(I).Tab.Color = RGB(TabColor(4) , TabColor(3) , TabColor(2))
                end if

            next

            '把各个Sheet里的背景纹理全部重置（涂满）
            Worksheets(I).Cells.Interior.Pattern = xlgray16
            Worksheets(I).Cells.Interior.PatternColor = 0
        end if
    next

    
    if Remark_All_Sheets = False then  '部分重置模式还是要重置已选定的表的
        for I = 1 to ActiveWindow.SelectedSheets.Count
            With ActiveWindow.SelectedSheets(I)
                .Tab.Color = RGB(0,0,0)
                .Cells.Interior.Pattern = xlgray16
                .Cells.Interior.PatternColor = 0
            End With
        Next
    end if


    '遍历已选定Sheet
    Edit_Sheet_Name = ""
    for SkewerFinger_Row = 1 to SkewerFinger_MaxRow()
        With Global_SkewerHeader.offset(SkewerFinger_Row,0)
            '判断行是否符合条件
            if .Cells(1).Text = X_CutName and .Cells(3).Text <> "" and .Cells(3).Text <> Edit_Sheet_Name then
                Edit_Sheet_Name = .Cells(3)
                Edit_Range_Address = .Cells(4)

                if activeworkbook.Worksheets(Edit_Sheet_Name).Tab.Color >0 then  '如果标签已经是黑色的了就跳过它

                    '开始设定
                    With activeworkbook.Worksheets(Edit_Sheet_Name)
                        .Activate  '设为当前编辑

                        '已选定Sheet标签涂成黑色
                        .Tab.Color = RGB(0,0,0)

                        '把已选定区域的纹理去掉
                        .range(Edit_Range_Address).Interior.Pattern = xlNone

                        .range(Edit_Range_Address).Select
                        ActiveWindow.FreezePanes = false '取消冻结
                        ActiveWindow.Zoom = True   '自动缩放显示

                        If ActiveWindow.Zoom > 150 then ActiveWindow.Zoom = 150     '如果Zoom过大，就缩小回去
                        If ActiveWindow.Zoom <45 and Selection.Rows.Count > 3 then   '最小缩放45%

                            '只考虑高度，超宽的情况暂不考虑，表太大了，应用情况多再加这个功能吧

                            ActiveWindow.Zoom <45
                            ActiveWindow.ScrollRow = 1  
                            ActiveWindow.ScrollColumn = 1    '先恢复到起始位置

                            Edit_Range_LastCell = .range(Edit_Range_Address).cells(.range(Edit_Range_Address).Cells.Count)
                            
                            For I = 1 To .range(Edit_Range_Address).Rows.Count / 3

                                Split_Range_Height = 0
                                Split_Range_Height = Split_Range_Height + .Range(.Cells(1) , .Cells(1).offset(I,0)).Height '加上半高度
                                Split_Range_Height = Split_Range_Height + .Range(Edit_Range_LastCell , Edit_Range_LastCell.offset(-2 * I,0)) '加下半高度

                                If ActiveWindow.Height / Split_Range_Height < 0.5 then Exit For

                            Next
                            
                            I = I - 1

                            .Cells(1).offset(I,0).Select
                            ActiveWindow.FreezePanes = True   '冻结 I 行

                            ActiveWindow.ScrollRow = Edit_Range_LastCell.offset(-2 * I,0) . Row '滚动到 -2I
                        End If
                    End With
                end if
            end if
        End With
    next
    

    ActiveWorkbook.Saved = True  '标记成不用保存，保存很麻烦的~~~~~~~~~~~~~~~
End Sub

'Sheet,Range 设定完成
Sub Skewer_RangeCommit(optional Cal As XlCalculation)
    Dim I as Integer
    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    for I = 1 to activeworkbook.names.count  '先判断一下是不是已经记录了颜色
        if activeworkbook.names(I).name = "Skewer_SheetColor" then exit for
    next
    If I > activeworkbook.names.count then
        msgbox("当前文件未经处理")
        exit sub
    End If

    Call Skewer_FileWalker_Next
End Sub

'开始导入
Sub Skewer_LoadRanges(optional Cal As XlCalculation)
    Dim I as Long , J as Long
    Dim Checked_CutName as String
    Dim SkewerFinger_Row as Long

    Dim X_CutName as String , X_Label as String , X_Sheet as String , X_Range as String

    Dim X_Workbook as String
    Dim Range_To_Load as Range

    '刷新 Global_SkewerWorkbook 和 Global_SkewerHeader
    If Not Find_SkewerWorkbook() then Exit Sub

    '验证当前workbook Sheet
    If activeworkbook.Name <> Global_SkewerWorkbook.Name then
        msgbox("当前文件不是 SkewerWorkbook")
        exit Sub
    end if

    If ActiveSheet.Name <> Global_SkewerHeader.Worksheet.Name then
        msgbox("当前Sheet不是 SkewerHeader.Worksheet")
        exit Sub
    end if

    '检测相关联文件是否有编辑情况，关闭所有相关联的文件
    Checked_CutName = ""
    for I = 1 to SkewerFinger_MaxRow()
        With Global_SkewerHeader.offset(I,0)
            if .cells(1).Text <> Checked_CutName then
                for J=1 to workbooks.Count
                    if right(.cells(1).Text,len(workbooks(J).Name) + 1) = "\" & workbooks(J).Name then
                        msgbox("检测到文件名冲突，需要关闭")
                        workbooks(J).Close
                    end if

                    if right(.cells(1).Text,len(workbooks(J).Name) + 1) = "\" & workbooks(J).Name then
                        msgbox("文件关闭失败，程序终止")
                        Exit Sub
                    end if
                next
                Checked_CutName = .cells(1).Text
        end with
    next
    'SkewerFinger_Row FingerLen 遍历
    for SkewerFinger_Row = SkewerFinger_MaxRow() to 1 step -1
        With Global_SkewerHeader.offset(I,0)
            '读取 删除行
            X_CutName = .Cells(1).Text
            X_Label = .Cells(2).Text
            X_Sheet = .Cells(3).Text
            X_Range = .Cells(4).Text
            .EntireRow.Delete  '先删除了再说

            '检测是不是第一行
            If Not( _
                X_CutName = .Cells(1).offset(-1,0).Text and _
                X_Label = .Cells(2).offset(-1,0).Text and _
                X_Sheet = .Cells(3).offset(-1,0).Text and _
                X_Range = .Cells(4).offset(-1,0).Text _
            ) then  '如果是该Sheet的第一行就开始导入

                X_Workbook = ""

                for J = 1 to workbooks.Count
                    if right(workbooks(I).FullName , len(X_CutName)) = X_CutName then
                        X_Workbook = workbooks(I).Name
                        exit for
                    end if
                next

                '检测文件是否已经打开，没有就打开
                if X_Workbook = "" then
                    X_Workbook = (application.workbooks.open Filename:= Global_SkewerWorkbook.Path & X_CutName, ReadOnly:=True).Name  '打开文件 记录名字
                    Global_SkewerWorkbook.Activate '当前文件切换回 SkewerWorkbook
                End If

                Set Range_To_Load = Workbooks(X_Workbook).Worksheets(X_Sheet).Range(X_Range)

                'FingerLen 按目标Range扩充
                .Worksheet.Range(.cells(1) , .cells(1).offset(Range_To_Load.Height - 1 , 0 )) _  
                    .EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

                '复制数据
                Range_To_Load.Copy
                .cells(1).offset(0,5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                .cells(1).offset(0,5).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

            End If
        end with
    next

    msgbox("数据复制完成")

End Sub