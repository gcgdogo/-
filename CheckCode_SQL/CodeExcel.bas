Attribute VB_Name = "CodeExcel"
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
ToolA = CodeRnd(i Xor Code)
ToolB = Not ToolA
A = A Xor (Y And ToolA)
B = B Xor (Y And ToolB)
Next
CodeStr = CLng(CodeRnd(A)) * 65536 + CodeRnd(B)
End Function
Function CodeRge(X As Range, Optional Code As Integer = 1) As Long
Dim Y As Range, A As String, i As Integer, Out As Long
i = 0
A = ""
Out = 0
For Each Y In X.Cells
i = i Xor Code
Out = Out Xor CodeStr(CStr(Y), i)
Next
CodeRge = Out
End Function
Function CodeHex(X As Range, Optional Code As Integer = 1) As String
CodeHex = Right("00000000" + Hex(CodeRge(X, Code)), 8)
End Function

