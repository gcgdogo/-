Option Explicit
Dim Tar_Workbook as Workbook
Dim Sheet_Table as Worksheet , Sheet_MarkData as Worksheet

Dim RM_Data_Rows as Long  'RM_Read可以使用的最大行数放在全局里面，读取的时候还要用呢

'连接数据库
'从数据库中读取区域标记情况


'定义Range_Marker获取列名的方式，方便日后读取
function RM_Read(x_Key as String , x_Row as Long , optional Force_Update as boolean = False) as Variant

    'RM_Data(Key轴)(Row轴)
    '标题存在第 0 行
    '刷新数据方法  :  RM_Read("" , 0 , Force_Update := True)

    Static RM_Data()() as Variant , RM_Data_Keys as Integer , RM_Data_keyscount Update_Time as double
    Dim I_Key as Long , I_Row as Long
    

    '根据条件，重新刷新一次
    if timer() - Update_Time > 5 then Force_Update = True  '超过5秒就自动重新读取一次
    if Force_Update then
        RM_Data_Keys = Sheet_MarkData.UsedRange.Columns.Count
        RM_Data_Rows = Sheet_MarkData.UsedRange.Rows.Count
        ReDim RM_Data(RM_Data_Keys + 1)(RM_Data_Rows + 1)
        for I_Key = 1 to RM_Data_Keys
            for I_Row = 1 to RM_Data_Rows
                RM_Data(I_Key)(I_Row - 1) = Sheet_MarkData.UsedRange.Columns(I_Key).cells(I_Row)   '标题存在第 0 行，行数减一
            next
        next
        Update_Time = timer()
    end if


    'x_Key = "" 意味着不用读取，用于刷新的情况
    if x_Key = "" then exit function 

    '检查要提取的行号
    if x_Row > RM_Data_Rows then
        msgbox("function RM_Read : x_Row > RM_Data_Rows")
        exit function
    end if

    '查找列信息
    for I_Key = 1 to RM_Data_Keys
        if RM_Data(I_Key)(0) = x_Key then
            RM_Read = RM_Data(I_Key)(x_Row)
            exit function
        end if
    next

    msgbox("function RM_Read:[Not Fond] x_Key = " & str(x_Key))  '没找到对应列

end function


'调整区域空间，添加辅助行列
sub Adjust_Range_AddRC()
    Dim re_testRow as new RegExp
    Dim re_testCol as new RegExp
    Dim Top_MarkRows_Count as Long , Left_MarkCols_Count as Long

    Dim ori_DataRange as Range , new_DataRange as Range   '记录原始数据位置和移动后的数据位置
    Dim combine_Mark_And_Data_Range as Range  '记录一下最大范围的Range

    Dim I as Long , I_Cell as Range

    Dim Temp_Address as String  '读取列标题字母的时候临时用一下

    call RM_Read("" , 0 , Force_Update := True) '现刷新一下数据冷静一下

    re_testCol.Pattern = "[a-zA-Z]{1,3}:[a-zA-Z]{1,3}"
    re_testRow.Pattern = "[0-9]{1,4}:[0-9]{1,4}"
    for I = 1 to RM_Data_Rows  '遍历检测需要辅助标记的行列数量
        if re_testCol.Test(RM_Read("区域",I)) then
            if RM_Read("标记位置",I) > Top_MarkRows_Count then Top_MarkRows_Count = RM_Read("标记位置",I)
        end if
        if re_testRow.Test(RM_Read("区域",I)) then
            if RM_Read("标记位置",I) > Left_MarkCols_Count then Left_MarkCols_Count = RM_Read("标记位置",I)
        end if
    next

    '读取数据区域
    for I = 1 to RM_Data_Rows   '遍历读取数据位置
        if RM_Read("类型",I) = "文件信息_表格区域" then
            set ori_DataRange = Sheet_Table.range(RM_Read("区域",I))
            set new_DataRange = Sheet_Table.range(RM_Read("区域",I))
            exit for
        end if
    next

    '先插入一行一列把原列好位置标记上
    new_DataRange.cells(1).entirerow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    for each I_Cell in new_DataRange.rows(1).cells
        Temp_Address = I_Cell.entirecolumn.address(ColumnAbsolute := False) '获取一个不带$的地址
        Temp_Address = left(int(len(Temp_Address)/2)) '除二，截断一下
        I_Cell.value = Temp_Address
    next
    set new_DataRange = new_DataRange.offset(1,0) '移动一下，作为标记

    '再插入一列把原行号标记上
    new_DataRange.cells(1).entirecolumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    for each I_Cell in new_DataRange.columns(1).cells
        I_Cell.value = I_Cell.row - 1  '行号好办，因为已经下移一行了，所以要减一
    next
    set new_DataRange = new_DataRange.offset(0,1) '更新数据区域位置

    '插入全部标记所需位置
    
    ori_DataRange.cells(1).entirecolumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove  '插入最左侧的全局标记
    set new_DataRange = new_DataRange.offset(0,1) '移动一下，作为标记

    for I = 1 to Top_MarkRows_Count
        new_DataRange.cells(1).entirerow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove   '顶部标记区域
        set new_DataRange = new_DataRange.offset(1,0) '移动一下，作为标记
    next

    for I = 1 to Left_MarkCols_Count
        new_DataRange.cells(1).entirecolumn.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove   '左侧标记区域
        set new_DataRange = new_DataRange.offset(0,1) '移动一下，作为标记
    next

    '计算，保存合并后的最大区域
    set combine_Mark_And_Data_Range = Sheet_Table.Range( _
        ori_DataRange.cells(1).address & ":" & _
        new_DataRange.cells(new_DataRange .cells.count).address)
    
    '对新增的行列统统格式化
    for I = ori_DataRange.cells(1).row to new_DataRange.cells(1).row - 1
        call RM_Format_UnusedMarkRange(Sheet_Table.rows(I))  '先整体涂色
        call RM_Format_MarkRange(Intersect(Sheet_Table.rows(I),combine_Mark_And_Data_Range))  '对标注范围内的区域涂色
    next
    for I = ori_DataRange.cells(1).column to new_DataRange.cells(1).column - 1
        call RM_Format_UnusedMarkRange(Sheet_Table.columns(I))
        call RM_Format_MarkRange(Intersect(Sheet_Table.columns(I),combine_Mark_And_Data_Range))
    next

    '对辅助坐标区域和首列全局标记格式化
    call RM_Format_OriginalAddress(combine_Mark_And_Data_Range.rows(1))
    call RM_Format_OriginalAddress(combine_Mark_And_Data_Range.columns(2))

    call RM_Format_MarkRange(combine_Mark_And_Data_Range.columns(1))
end sub

