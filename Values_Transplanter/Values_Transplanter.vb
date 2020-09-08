option explicit

dim xValues(1000) as variant
dim xValuesCount as Integer

sub Values_Load(optional control As IRibbonControl)
    dim I_Cell as Range
    xValuesCount = 0
    for each I_Cell in selection.cells
        if I_Cell.address = I_Cell.MergeArea.cells(1).address then   '只计算合并单元格首格
            xValuesCount = xValuesCount + 1
            xValues(xValuesCount) = I_Cell
        end if
    next
    Msgbox("数据读取完毕：" & vbcrlf & "Range = " & selection.address & vbcrlf & "xValuesCount = " & cstr(xValuesCount))
end sub

sub Values_Write(optional control As IRibbonControl)
    dim I_Cell as Range , I_Cell_Count as Integer
    I_Cell_Count = 0
    for each I_Cell in selection.cells
        if I_Cell.address = I_Cell.MergeArea.cells(1).address then   '只计算合并单元格首格
            I_Cell_Count = I_Cell_Count + 1
        end if
    next

    if I_Cell_Count <> xValuesCount then
        msgbox("警告：区域大小不一致 xValuesCount = " & cstr(xValuesCount) & " , I_Cell_Count = " & cstr(I_Cell_Count))
        exit sub
    end if

    I_Cell_Count = 0
    for each I_Cell in selection.cells
        if I_Cell.address = I_Cell.MergeArea.cells(1).address then   '只计算合并单元格首格
            I_Cell_Count = I_Cell_Count + 1
            Call Try_WriteValue(I_Cell , xValues(I_Cell_Count) )
        end if
    next
    Msgbox("数据写入完毕：" & vbcrlf & "xValuesCount = " & cstr(xValuesCount))
end sub

sub Try_WriteValue(X_Cell as Range , X_Value as variant)
    On Error GoTo LineFail
    X_Cell = X_Value
LineFail:
end sub