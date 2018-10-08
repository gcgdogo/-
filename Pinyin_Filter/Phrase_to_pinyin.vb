Option Explicit

Function Phrase_to_pinyin(P as String) as String
	Dim I as Integer , J as Integer 
	Dim Pinyin as String,Pinyin_Temp as String
	Dim RegExp_Replace as new RegExp
	Dim RegExp_Match as new RegExp
	Dim R_Matches as MatchCollection
	
	Pinyin=","
	
	RegExp_Replace.IgnoreCase = True
	RegExp_Replace.Global = True
	RegExp_Replace.Pattern=","

	RegExp_Match.IgnoreCase=True
	RegExp_Match.Global=True
	RegExp_Match.Pattern="[^,]*,"

	For I=1 To Len(P)
		set R_Matches = RegExp_Match.Execute(Character_to_Pinyin(Mid(P,I,1)))
		Pinyin_Temp=""
		'MatchCollection 的集合从0开始到n-1
		For J=0 To R_Matches.Count-1
			Pinyin_Temp = Pinyin_Temp & RegExp_Replace.Replace(Pinyin,R_Matches(J).Value)
		Next
		Pinyin=Pinyin_Temp
		'Msgbox("I=" & I & " : " & Pinyin)
	Next
	Phrase_to_pinyin=Pinyin
End Function