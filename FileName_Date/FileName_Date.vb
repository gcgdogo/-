Option Compare Database
Option Explicit

'附加值在另一个函数中单独计算,附加值变量在模块中定义
'正常匹配附加在整数位,剩余字符按照1/60附加
Dim Append_Value as Double
Dim Append_Text as String

Function FileName_Date(FileName As String) As Date
	Dim R_Exp as new RegExp , R_SubMatchs as SubMatches
	Dim X_Year as Integer , X_Month as Integer , X_Day as Integer

	'模块级变量重置
	Append_Value=0
	Append_Text=""

    R_Exp.IgnoreCase=True
    R_Exp.Global=False

    '识别范围 .xls .xlsx 文件名后缀变化则需要修改 支持日期形式 [1234.56.78,1234-56-78,12-3-4]
    R_Exp.Pattern="([0-9]{2,4})[-\.]([0-9]{1,2})[-\.]([0-9]{1,2})(.*)\.xls.?$"

    '正则表达式识别失败就启用旧版方案
    If R_Exp.Test(FileName)=False Then
    	FileName_Date=FileName_Date_PlanB(FileName)
    	Exit Function
	End If

	'没问题就继续
	Set R_SubMatchs = R_Exp.Execute(FileName).Item(0).SubMatches

	X_Year=CInt(R_SubMatchs(0))
	'年份小于1800就按照20xx计算
	If X_Year<1800 Then X_Year= 2000 + (X_Year mod 100)

	X_Month=CInt(R_SubMatchs(1))
	X_Day=CInt(R_SubMatchs(2))

	Append_Text = R_SubMatchs(3)
	Call Append_Value_Cal

	FileName_Date = DateSerial(X_Year , X_Month , X_Day) + Append_Value/24
End Function

'需要识别的后缀文本列表在这里
Private Sub Append_Value_Cal()
	Call Append_Value_Phrase( ".Lang2XL" , 0 )
	Call Append_Value_Phrase( "#Lang2XL" , 0 )
	Call Append_Value_Phrase( "岗级" , 1 )
	Call Append_Value_Phrase( "v0" , 0.5 )
	Call Append_Value_Phrase( "v1" , 1 )
	Call Append_Value_Phrase( "v2" , 2 )
	Call Append_Value_Phrase( "v3" , 3 )
	Call Append_Value_Phrase( "v4" , 4 )
	Call Append_Value_Phrase( "v5" , 5 )
	Call Append_Value_Phrase( "v6" , 6 )
	Call Append_Value_Phrase( "v7" , 7 )
	Call Append_Value_Phrase( "v8" , 8 )
	Call Append_Value_Phrase( "v9" , 9 )
	Call Append_Value_Phrase( "A" , 1 )
	Call Append_Value_Phrase( "B" , 2 )
	Call Append_Value_Phrase( "C" , 3 )
	Call Append_Value_Phrase( "D" , 4 )
	Call Append_Value_Phrase( "E" , 5 )
	Call Append_Value_Phrase( "F" , 6 )
	Call Append_Value_Phrase( "G" , 7 )
	Call Append_Value_Phrase( "H" , 8 )

	'最后将剩余字符按1/60计算
	Append_Value = Append_Value + Len(Append_Text)/60
End Sub

'将已识别的文本剔除,处理后文本作为返回值返回
Private Sub Append_Value_Phrase(Check_Text as String , Check_Value as Double)
	Dim R_Exp as new RegExp
    R_Exp.IgnoreCase=True
    R_Exp.Global=True

    '为的特殊字符增加转义字符 "([\$\(\)\[\]\{\}\.\+\?])"
    R_Exp.Pattern="([\$\(\)\[\]\{\}\.\+\?])"
    R_Exp.Pattern= R_Exp.Replace(Check_Text,"\$1")

    If R_Exp.Test(Append_Text)=False Then Exit Sub

    Append_Value = Append_Value + Check_Value
    Append_Text = R_Exp.Replace(Append_Text , "")
    
End Sub

'旧版算法作为备用方案
Private Function FileName_Date_PlanB(FileName As String) As Date
	Dim I As Integer, J As Integer, X As String, A(20) As Integer
	J = 1
	For J = 1 To 19
	    A(J) = 0
	Next
	J = 1
	For I = 1 To Len(FileName)
	    X = Mid(FileName, I, 1)
	    If X >= "0" And X <= "9" Then
	        A(J) = A(J) * 10 + CInt(X)
	    Else
	        If A(J) > 0 Then J = J + 1
	    End If
	Next
	FileName_Date_PlanB = DateSerial(A(1), A(2), A(3))
End Function

Private Sub FileName_Date_test()
MsgBox (CStr(FileName_Date("asderfe 1859 asc d10  sd4 asdf568")))
End Sub