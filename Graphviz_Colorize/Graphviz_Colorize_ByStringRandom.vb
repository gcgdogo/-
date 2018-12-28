Option Explicit

Public Function GColorize_String (Seed_A as String , Optional Seed_B as String = "" ) as String
    Dim RGB_A as Long , RGB_B as Long , I as Long
    Dim GradualStages as Long

    If Seed_B = "" Then
        GColorize_String = "#" & Right("000000" & Hex(Get_Color(Get_SeedCode(Seed_A))) , 6)
        'msgbox(Seed_A & Get_SeedCode(Seed_A) & Hex(Get_Color(Get_SeedCode(Seed_A))))
    Else
        RGB_A = Get_Color(Get_SeedCode(Seed_A))
        RGB_B = Get_Color(Get_SeedCode(Seed_B))

        If RGB_A = RGB_B Then
            GColorize_String = "#" & Right("000000" & Hex(Get_Color(Get_SeedCode(Seed_A))) , 6)
            Exit Function
        End If

        GradualStages = Get_GradualStages( RGB_A , RGB_B )
        GColorize_String=""

        For I = 1 To GradualStages
            GColorize_String = GColorize_String & "#" & Right("000000" & Hex(RGB_Gradual(RGB_A , RGB_B , CDbl(I-1) / (GradualStages - 1))) , 6)
            '附加一个比例数来将颜色纵向排列, 比例数格式化为8位小数
            If I = 1 Then GColorize_String = GColorize_String & ";" & Format( 1.0 / GradualStages , "0.00000000" )
            If I < GradualStages Then GColorize_String = GColorize_String & ":"
        Next
    End If

End Function

Private Function Get_GradualStages(RGB_A as Long , RGB_B as Long) as Long
    Dim RGB_A_1 as Double , RGB_A_2 as Double ,  RGB_A_3 as Double
    Dim RGB_B_1 as Double , RGB_B_2 as Double ,  RGB_B_3 as Double
    Dim Color_Steps as Integer

    '设置颜色级差
    Color_Steps = 10

    RGB_A_1 = Int(RGB_A / 256 / 256)
    RGB_A_2 = Int(RGB_A / 256) mod 256
    RGB_A_3 = RGB_A mod 256

    RGB_B_1 = Int(RGB_B / 256 / 256)
    RGB_B_2 = Int(RGB_B / 256) mod 256
    RGB_B_3 = RGB_B mod 256
    
    Get_GradualStages = 2

    If ( Abs(RGB_A_1 - RGB_B_1) / Color_Steps ) > Get_GradualStages Then Get_GradualStages = ( Abs(RGB_A_1 - RGB_B_1) / Color_Steps )
    If ( Abs(RGB_A_2 - RGB_B_2) / Color_Steps ) > Get_GradualStages Then Get_GradualStages = ( Abs(RGB_A_2 - RGB_B_2) / Color_Steps )
    If ( Abs(RGB_A_3 - RGB_B_3) / Color_Steps ) > Get_GradualStages Then Get_GradualStages = ( Abs(RGB_A_3 - RGB_B_3) / Color_Steps )

End Function

Private Function RGB_Gradual(RGB_A as Long , RGB_B as Long , Percent as Double) as Long
    Dim RGB_A_1 as Double , RGB_A_2 as Double ,  RGB_A_3 as Double
    Dim RGB_B_1 as Double , RGB_B_2 as Double ,  RGB_B_3 as Double
    Dim RGB_1 as Double , RGB_2 as Double ,  RGB_3 as Double

    RGB_A_1 = Int(RGB_A / 256 / 256)
    RGB_A_2 = Int(RGB_A / 256) mod 256
    RGB_A_3 = RGB_A mod 256

    RGB_B_1 = Int(RGB_B / 256 / 256)
    RGB_B_2 = Int(RGB_B / 256) mod 256
    RGB_B_3 = RGB_B mod 256

    RGB_1 = Int(RGB_A_1 * (1 - Percent) + RGB_B_1 * Percent)
    RGB_2 = Int(RGB_A_2 * (1 - Percent) + RGB_B_2 * Percent)
    RGB_3 = Int(RGB_A_3 * (1 - Percent) + RGB_B_3 * Percent)

    RGB_Gradual = RGB_1 * 256 * 256 + RGB_2 * 256 + RGB_3

End Function

Private Function Get_Color(Seed_Code as Double) as Long
    Dim Get_Red as Integer , Get_Green as Integer , Get_Blue as Integer
    Dim I as Integer

    '基础颜色列表先以50个为最大限额
    Dim Src_Count as Integer , Src_Code(50) as Double , Src_Red(50) as Double , Src_Green(50) as Double , Src_Blue(50) as Double

    Dim Upper_Code as Double , Upper_Red as Double , Upper_Green as Double , Upper_Blue as Double
    Dim Lower_Code as Double , Lower_Red as Double , Lower_Green as Double , Lower_Blue as Double
    Dim Upper_Factor as Double , Lower_Factor as Double
    Src_Count=0
    '基础颜色列表:
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0 : Src_Red(Src_Count) = 137 : Src_Green(Src_Count) = 106 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.0625 : Src_Red(Src_Count) = 100 : Src_Green(Src_Count) = 125 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.125 : Src_Red(Src_Count) = 22 : Src_Green(Src_Count) = 132 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.1875 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 135 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.25 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 134 : Src_Blue(Src_Count) = 20
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.3125 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 134 : Src_Blue(Src_Count) = 65
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.375 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 131 : Src_Blue(Src_Count) = 130
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.4375 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 106 : Src_Blue(Src_Count) = 176
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.5 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 42 : Src_Blue(Src_Count) = 223
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.5625 : Src_Red(Src_Count) = 0 : Src_Green(Src_Count) = 0 : Src_Blue(Src_Count) = 235
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.625 : Src_Red(Src_Count) = 74 : Src_Green(Src_Count) = 0 : Src_Blue(Src_Count) = 223
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.6875 : Src_Red(Src_Count) = 147 : Src_Green(Src_Count) = 0 : Src_Blue(Src_Count) = 199
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.75 : Src_Red(Src_Count) = 174 : Src_Green(Src_Count) = 0 : Src_Blue(Src_Count) = 164
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.8125 : Src_Red(Src_Count) = 179 : Src_Green(Src_Count) = 0 : Src_Blue(Src_Count) = 65
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.875 : Src_Red(Src_Count) = 178 : Src_Green(Src_Count) = 0 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 0.9375 : Src_Red(Src_Count) = 170 : Src_Green(Src_Count) = 42 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 1 : Src_Red(Src_Count) = 137 : Src_Green(Src_Count) = 106 : Src_Blue(Src_Count) = 0
        Src_Count = Src_Count + 1 : Src_Code(Src_Count) = 1.0625 : Src_Red(Src_Count) = 100 : Src_Green(Src_Count) = 125 : Src_Blue(Src_Count) = 0


    '寻找相邻项目
    Upper_Code=2
    Lower_Code=-1
    For I = 1 to Src_Count
        If Src_Code(I) > Seed_Code Then
            If Src_Code(I) < Upper_Code Then 
                Upper_Code = Src_Code(I)
                Upper_Red = Src_Red(I)
                Upper_Green = Src_Green(I)
                Upper_Blue = Src_Blue(I)
            End If
        Else
            If Src_Code(I) > Lower_Code Then 
                Lower_Code = Src_Code(I)
                Lower_Red = Src_Red(I)
                Lower_Green = Src_Green(I)
                Lower_Blue = Src_Blue(I)
            End If
        End If
    Next

    '差值计算对应颜色
    Upper_Factor = (Seed_Code - Lower_Code) / (Upper_Code - Lower_Code)
    Lower_Factor = (Upper_Code - Seed_Code) / (Upper_Code - Lower_Code)
    Get_Red = Upper_Red * Upper_Factor + Lower_Red * Lower_Factor
    Get_Green = Upper_Green * Upper_Factor + Lower_Green * Lower_Factor
    Get_Blue = Upper_Blue * Upper_Factor + Lower_Blue * Lower_Factor

    'Graphviz RGB 顺序 与 VB中RGB顺序颠倒 需要倒着输入
    Get_Color = RGB(Get_Blue , Get_Green , Get_Red)

End Function


'Graph_Jdg中以特殊字符之前的文本为颜色计算基础,便于条件信息染色
Private Function Get_SeedCode(Seed as String) as Double
    Dim Asc_Combine as Integer
    Dim I as Integer
    Dim Seed_Char as String

    Asc_Combine = 0
    For I = 1 To Len(Seed)
        Seed_Char = Mid(Seed , I , 1)
        '特殊字符判定
        If Seed_Char = "=" or Seed_Char = "<" or Seed_Char = ">" or Seed_Char = "[" or Seed_Char = "]" Then Exit For
        Asc_Combine = Asc_Combine Xor CodeRnd(Asc(Seed_Char))
    Next

    Get_SeedCode= Abs( CDbl(CodeRnd(Asc_Combine)) ) / 32768

End Function

Private Function CodeRnd(X As Integer) As Integer
    CodeRnd = ((Sin(X) + 1) * 268435456) Mod 65536 - 32768
End Function
