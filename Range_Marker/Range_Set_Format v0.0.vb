Option Explicit
'为避免写错类型代码，直接用单独定义函数的形式表示不同的格式

sub RM_Format_UnusedMarkRange(x_Range as range)  '未使用的标记行
    call Set_Format_NoInsideBorder
    with x_Range
        .Interior.Pattern = xlNone
        .Interior.color = rgb(80,80,80)  '灰色背景
    end with
end sub

sub RM_Format_OriginalAddress(x_Range as range)  '用来表示原地址的区域
    call Set_Format_NormalBalckBorder(x_Range)
    with x_Range
        .Interior.Pattern = xlNone
        .Interior.color = rgb(140,140,140)  '灰色背景
        .Font.Color = rgb(255,255,255)
        .Font.Bold = True
        .Font.Size = 11   '白色加粗文字
        .Font.Name = "宋体"
    end with
end sub

sub RM_Format_MarkRange(x_Range as range)  '用来表示进行标记的区域
    call Set_Format_NormalBalckBorder(x_Range)
    with x_Range
        .Interior.Pattern = xlNone
        .Interior.color = rgb(200,200,200)  '灰色背景
        .Font.Color = rgb(255,255,255)
        .Font.Bold = False
        .Font.Size = 11   '白色加粗文字
        .Font.Name = "宋体"
    end with
end sub

'正常的内外黑边框
private sub Set_Format_NormalBalckBorder(x_Range as range)
    with x_Range
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End with
end sub

'内部无边框
private sub Set_Format_NoInsideBorder(x_Range as range)
    with x_Range
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
    End with
end sub