Option Compare Database
Option Explicit

'附加值在另一个函数中单独计算,附加值变量在模块中定义
'正常匹配附加在整数位,剩余字符按照1/60附加
Dim Append_Value as Double

Function FileName_Date(FileName As String) As Date
	Dim R_Exp as new RegExp , R_SubMatchs as SubMatches
	Dim X_Year as Integer , X_Month as Integer , X_Day as Integer
	Dim Append_Text as String

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



End Function

'将已识别的文本剔除,处理后文本作为返回值返回
Private Function Append_Value_Cal(Append_Text as String , Check_Text as String , Check_Value as Double) as String
	Dim R_Exp as new RegExp
    R_Exp.IgnoreCase=True
    R_Exp.Global=True
    R_Exp.Pattern=Check_Text
    If R_Exp.Test(Append_Text)=False Then Exit Function


    
End Function

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