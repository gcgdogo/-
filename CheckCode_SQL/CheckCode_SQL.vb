Option Explicit

Function CodeRnd(X As Integer) As Integer
	CodeRnd = ((Sin(X) + 1) * 268435456) Mod 65536 - 32768
End Function

Function CodeStr(X As String, Optional Code As Integer = 0) As Long
	Dim i As Integer
	Dim ToolA As Integer, ToolB As Integer, A As Integer, B As Integer, Y As Integer
	A = 0
	B = 0
	For i = 1 To Len(X)
		Y = CodeRnd(Asc(Mid(X, i, 1)))
		ToolA = CodeRnd(CodeRnd(i) Xor Code)
		ToolB = Not ToolA
		A = A Xor (Y And ToolA)
		B = B Xor (Y And ToolB)
	Next
	A = CodeRnd(A)
	B = CodeRnd(B)
	CodeStr = CLng(A) * 65536 + 32768 + B
End Function


Function CheckCode_SQL_Lng (X as String) As Long
	Dim ADO_rs as new ADODB.recordset
	Dim I as integer , Test_String as String
	Dim Out as Long
	Out = 0
	ADO_rs.ActiveConnection = CurrentProject.Connection
	ADO_rs.CursorType=adOpenStatic
	ADO_rs.LockType=adLockReadOnly
	ADO_rs.Source=X
	ADO_rs.Open

	Do While ADO_rs.EOF = False
		Test_String= ""
		For I=0 To ADO_rs.Fields.Count - 1
			Test_String=Test_String & CStr(nz(ADO_rs(ADO_rs.Fields(I).Name),""))
		Next
		Out =Out Xor CodeStr(Test_String)
		ADO_rs.MoveNext
	Loop
	CheckCode_SQL_Lng = Out
End Function

Function CheckCode_SQL_Hex(X As String) As String
	CheckCode_SQL_Hex = Right("00000000" + Hex(CheckCode_SQL_Lng(X)), 8)
End Function
