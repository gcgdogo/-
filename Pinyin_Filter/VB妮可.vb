' VB妮可
' 推荐于2016-01-14
' 你这个汉字转拼音的函数是常用汉字的,取的是ASCII从-20319到-10247 的汉字

' 你的孢，孚都是这个范围之外的..

' 解决方法只能再找更全的码表呵..比如包含也有的GBK文字的

' 补充:我找到的另类的解决方法:

' ====================================
' '模块:
Option Explicit

Private Const IME_ESC_MAX_KEY = &H1005
Private Const IME_ESC_IME_NAME = &H1006
Private Const GCL_REVERSECONVERSION = &H2
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Private Declare Function ImmEscape Lib "imm32.dll" Alias "ImmEscapeA" (ByVal hkl As Long, ByVal himc As Long, ByVal un As Long, lpv As Any) As Long
Private Declare Function ImmGetConversionList Lib "imm32.dll" Alias "ImmGetConversionListA" (ByVal hkl As Long, ByVal himc As Long, ByVal lpsz As String, lpCandidateList As Any, ByVal dwBufLen As Long, ByVal uFlag As Long) As Long
Private Declare Function IsDBCSLeadByte Lib "kernel32" (ByVal bTestChar As Byte) As Long

Public Function GetChineseSpell(ByVal CHINESE As String, Optional PYTYPE As Integer = 0, Optional Delimiter As String = " ") As String

If Len(Trim(CHINESE)) > 0 Then
	Dim i As Long
	Dim s As String
	s = Space(255)
	Dim IMEInstalled As Boolean
	Dim j As Long
	Dim a() As Long

	ReDim a(255) As Long
	j = GetKeyboardLayoutList(255, a(LBound(a)))

	For i = LBound(a) To LBound(a) + j - 1
		If ImmEscape(a(i), 0, IME_ESC_IME_NAME, ByVal s) Then
			If Trim("微软拼音输入法") = Replace(Trim(s), Chr(0), "") Then
				IMEInstalled = True
			Exit For
		End If
		End If
	Next i
	If IMEInstalled Then
	CHINESE = Trim(CHINESE)
	Dim sChar As String
	Dim Buffer0() As Byte
	Dim bBuffer0() As Byte
	Dim bBuffer() As Byte
	Dim k As Long
	Dim l As Long
	Dim m As Long
	For j = 0 To Len(CHINESE) - 1
		sChar = Mid(CHINESE, j + 1, 1)
		' If Not InStr("《》，。/？、][{}“”‘’；：！·〈〉「」『』｜〖〗【】（）〔〕｛｝…—.,""'';:?/\!", sChar) > 0 Then
		Buffer0 = StrConv(sChar, vbFromUnicode)
		If IsDBCSLeadByte(Buffer0(0)) Then
			k = ImmEscape(a(i), 0, IME_ESC_MAX_KEY, Null)
			If k Then
				l = ImmGetConversionList(a(i), 0, sChar, 0, 0, GCL_REVERSECONVERSION)
				If l Then
					s = Space(255)
					If ImmGetConversionList(a(i), 0, sChar, ByVal s, l, GCL_REVERSECONVERSION) Then

						bBuffer0 = StrConv(s, vbFromUnicode)
						ReDim bBuffer(k * 2 - 1)
						For m = bBuffer0(24) To bBuffer0(24) + k * 2 - 1
							bBuffer(m - bBuffer0(24)) = bBuffer0(m)
						Next m
						sChar = Trim(StrConv(bBuffer, vbUnicode))
						If InStr(sChar, vbNullChar) Then
							sChar = Trim(Left(sChar, InStr(sChar, vbNullChar) - 1))
						End If
					End If
				End If

			End If
		End If
		' End If
		GetChineseSpell = GetChineseSpell & Switch(PYTYPE = 0, sChar, PYTYPE = 1, Left(sChar, Len(sChar) - 1), PYTYPE = 2, UCase(Left(sChar, 1))) & IIf(PYTYPE = 2, "", Delimiter) ''返回全拼
	Next j
	Else ''没安装“微软拼音输入法”,返回<未发现微软拼音>
	GetChineseSpell = "<未发现微软拼音>"
	End If
	Else
	GetChineseSpell = "" ''输入为空字符串
End If
End Function

'下面是窗体代码:
Private Sub Command1_Click()
Print GetChineseSpell("孢孚", 2)
End Sub

' ==============================
' 注意,
' 1.一定要系统安装的有微软拼音输入法,不然返回的是空格..
' 2.模块中没有带,或是说没完全带标点的处理过程,你应该自己在程序中处理或是修改模块
' 3.使用方法有3个参数,0是返回带单调的全拼,1是返回完整拼音,2是返回拼音首字母..

' 测试通过,VB妮可. 