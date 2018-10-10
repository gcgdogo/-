Option Explicit

Sub Range_to_Pinyin(control As IRibbonControl)
	Dim Range_A as Range , Range_B as Range
	Dim Range_Phrase as Range , Range_Pinyin as Range
	Dim X as Range
	Dim I as Long
	Dim Has_Value as Boolean

	If Selection.Columns.count>1 then
		set Range_A = Selection.Columns(1)
		set Range_B = Selection.Columns(2)
	End If

	If Selection.Areas.count>1 then
		set Range_A = Selection.Areas(1).Columns(1)
		set Range_A = Selection.Areas(2).Columns(1)
	End If

	If Range_A is Nothing or Range_B is Nothing then
		Msgbox("区域判定失败 : Range_A is Nothing or Range_B is Nothing")
		Exit Sub
	End If

	If Range_A.cells.count <> Range_B.cells.count then
		Msgbox("区域判定失败 : Range_A.cells.count <> Range_B.cells.count")
		Exit Sub
	End If

	Has_Value=False
	For each X in Range_A.cells
		if X.Text<>"" then Has_Value=true
	Next
	If Has_Value=false then
		set Range_Pinyin = Range_A
		set Range_Phrase = Range_B
	End If

	Has_Value=False
	For each X in Range_B.cells
		if X.Text<>"" then Has_Value=true
	Next
	If Has_Value=false then
		set Range_Pinyin = Range_B
		set Range_Phrase = Range_A
	End If

	If Range_Pinyin is Nothing or Range_Phrase is Nothing then
		Msgbox("区域判定失败 : 可能是 Has_Value=true , Range_Pinyin is Nothing or Range_Phrase is Nothing")
		Exit Sub
	End If

	For I=1 To Range_Phrase.cells.count
		Range_Pinyin.Cells(I)=Phrase_to_pinyin(Range_Phrase.Cells(I).Text)
	Next
End Sub