Option Compare Database
Option Explicit

Dim Str_Log as String

Function DateBlock_Formater( Source_Text as String , Time_Now as Variant , Divider_Year as String , Divider_Month as String , Divider_Connect as String , Divider_Divide as String) as String
    Dim Int_Year as Long , Int_Month as Long , Int_Day as Long , Int_Stock as Long
    Dim Flag_Connect as Boolean
    Dim I as Long , Char_I as String

    Dim Dates(101) as Variant , Dates_Count as Long
    Dim Date_Connect as Variant , Date_Connect_Target as Variant

    Dim DateBlock_String as String

    Str_Log = ""

    Int_Year = 0
    Int_Month = 0
    Int_Day = 0
    Int_Stock = 0

    Dates_Count = 0
    Flag_Connect = False

    Source_Text = Source_Text & Divider_Divide & " "   '在最后放个分隔符

    '循环获取日期
    For I = 1 to Len(Source_Text)
        Char_I = mid(Source_Text , I ,1 )

        Str_Log = Str_Log & Char_I & " : "
        if Flag_Connect = True then Str_Log = Str_Log & "[Connect] "

        if Dates_Count > 50 then exit for '暂时认为超过50个就错了

        If asc(Char_I) >= 48 and asc(Char_I) <= 57 then    '如果是数字
            Int_Stock = Int_Stock * 10 + Val(Char_I)
        else
            If Char_I = Divider_Year then
                Int_Year = Int_Stock
                Int_Stock = 0
                Char_I = "Nope" '防止触发其他判定 注意 "Nope" 与 "None" 的不同使用情况
                if Int_Year < 2000 then Int_Year = 2000 + Int_Year mod 100

                Str_Log = Str_Log & " Int_Year = " & Cstr(Int_Year)

                If Divider_Year = Divider_Month or Divider_Year = Divider_Connect or Divider_Year = Divider_Divide then Divider_Year = "None"  '如果其他分割也用，就只生效一次
            End If

            If Char_I = Divider_Month then
                Int_Month = Int_Stock
                Int_Stock = 0
                Char_I = "Nope" '防止触发其他判定 注意 "Nope" 与 "None" 的不同使用情况

                Str_Log = Str_Log & " Int_Month = " & Cstr(Int_Month)

                If Int_Year = 0 then  '如果年是空的就默认
                    If Int_Month <= month(Time_Now) + 1 then
                        Int_Year = year(Time_Now)
                    else
                        Int_Year = year(Time_Now) - 1
                    End If
                    Str_Log = Str_Log & " Int_Year = " & Cstr(Int_Year)
                End if
                If Divider_Month = Divider_Connect or Divider_Month = Divider_Divide then Divider_Month = "None"  '如果其他分割也用，就只生效一次
            End If

            If Char_I <> "" then '如果没清空就先把日子算了
                Int_Day = Int_Stock
                Int_Stock = 0

                If Int_Month = 0 then '如果月是空的就默认
                    If Int_Day <= day(Time_Now) + 10 then
                        Int_Month = month(Time_Now)
                    else
                        Int_Month = month(Time_Now) - 1
                    end If
                    Str_Log = Str_Log & " Int_Month = " & Cstr(Int_Month)
                end if

                If Int_Year = 0 then  '如果年是空的就默认
                    If Int_Month <= month(Time_Now) + 1 then
                        Int_Year = year(Time_Now)
                    else
                        Int_Year = year(Time_Now) - 1
                    End If
                    Str_Log = Str_Log & " Int_Year = " & Cstr(Int_Year)
                End if

                If Flag_Connect = True then
                    '循环递增到 Try_Date
                    If Try_Date(Int_Year , Int_Month , Int_Day) - Dates(Dates_Count) <= 30 then  '天数太多就不管了
                        Date_Connect = Dates(Dates_Count)
                        Date_Connect_Target = Try_Date(Int_Year , Int_Month , Int_Day)
                        Do While Date_Connect < Date_Connect_Target
                            Date_Connect = Date_Connect + 1
                            if Date_Connect = Date_Connect_Target or Holiday_Check(Date_Connect) = False  then   ' 检验是不是假期，段时间跳过假期，终点强制导入

                                Dates_Count = Dates_Count + 1
                                Dates(Dates_Count) = Date_Connect  '添加日期
                                Flag_Connect = False  '取消链接标记

                            end if

                            Str_Log = Str_Log + "<" & Cstr(Dates_Count) & ":" & Format( Dates(Dates_Count) , "yyyy-mm-dd") & ">"
                        Loop
                    End If
                else
                    Dates(Dates_Count + 1) = Try_Date(Int_Year , Int_Month , Int_Day)
                    if Dates(Dates_Count + 1) > 0 then Dates_Count = Dates_Count + 1

                    Str_Log = Str_Log + "<" & Cstr(Dates_Count) & ":" & Format( Dates(Dates_Count) , "yyyy-mm-dd") & ">"
                End If
            End if

            If Char_I = Divider_Connect then Flag_Connect = True

        end if

        Str_Log = Str_Log & vbCrLf

    Next


    For I = 1 to Dates_Count
        If I > 1 then
            DateBlock_String = DateBlock_String & vbCrLf '除第一行外加换行
            If Dates(I) - Dates(I-1) > 1 Then DateBlock_String = DateBlock_String & vbCrLf '如果日期不连续就加换行
        End if

        DateBlock_String = DateBlock_String & Format( Dates(I) , "yyyy-mm-dd")

    Next

    DateBlock_Formater = DateBlock_String

End Function


Function Try_Date(Int_Year as Long , Int_Month as Long , Int_Day as Long) as Variant
    Dim Try_Str as String
    On Error Goto Date_Failure
    Try_Str = Cstr(Int_Year) & "-" & Cstr(Int_Month) & "-" & Cstr(Int_Day)
    'Str_Log = Str_Log & "{Try:" & Try_Str & "}"
    Try_Date = DateValue(Try_Str)
    Exit Function
Date_Failure:
    Try_Date = 0
End Function

Function Get_Str_Log() as String
    Get_Str_Log = Str_Log
End Function