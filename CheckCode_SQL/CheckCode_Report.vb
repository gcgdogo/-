Diary_Add("Message",	"[" &
	Format(Dmin("时点ID","文件名_日期",
		"CheckCode=""" & 
Target_CheckCode
			 & """" &
		TimeID_Range_String
	),"00000000-000000")
	& ">>" &
	Format(Dmax("时点ID","文件名_日期",
		"CheckCode=""" & 
Target_CheckCode
			 & """" &
		TimeID_Range_String
	),"00000000-000000")
	&"] " & 
Target_CheckCode
)

Diary_Add("Message",
	"[" &
	Format(Dmin("时点ID","文件名_日期",
		"CheckCode=""" & 
Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000"))
			 & """" &
		" and 时点ID>" & Dmax("时点ID","文件名_日期","文件ID<900000")-600000000
	),"00000000-000000")
	& ">>" &
	Format(Dmax("时点ID","文件名_日期",
		"CheckCode=""" & 
Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000"))
			 & """" &
		" and 时点ID>" & Dmax("时点ID","文件名_日期","文件ID<900000")-600000000
	),"00000000-000000")
	&"] " & 
Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000"))
)



Diary_Add("Message",
	"[" &
	Format(Dmin("时点ID","文件名_日期",
		"CheckCode=""" & 
Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000 and CheckCode<>""" & 
	Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000")) & """"
))
			 & """" &
		" and 时点ID>" & Dmax("时点ID","文件名_日期","文件ID<900000")-600000000
	),"00000000-000000")
	& ">>" &
	Format(Dmax("时点ID","文件名_日期",
		"CheckCode=""" & 
Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000 and CheckCode<>""" & 
	Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000")) & """"
))
			 & """" &
		" and 时点ID>" & Dmax("时点ID","文件名_日期","文件ID<900000")-600000000
	),"00000000-000000")
	&"] " & 
Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000 and CheckCode<>""" & 
	Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000")) & """"
))
)


Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000 and CheckCode<>""" & 
	Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000")) & """"
))

Option Explicit

Function CheckCode_Report()
	Dim Target_CheckCode as String , TimeID_Range_String as String
	Dim Calling as Variant
	'先定位到前一个 CheckCode
	Target_CheckCode = Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000"))
	Target_CheckCode= Dfirst("CheckCode","文件名_日期", _
		"时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000 and CheckCode<>""" & Target_CheckCode & """") _
	)

	'时间范围条件文本
	TimeID_Range_String = " and 时点ID>" & Dmax("时点ID","文件名_日期","文件ID<900000")-600000000

	Calling=Diary_Add("Message", _
		"CheckCode_Report: [" & _
		Format(Dmin("时点ID","文件名_日期","CheckCode=""" & Target_CheckCode & """" & TimeID_Range_String),"00000000-000000") _
		& ">>" & _
		Format(Dmax("时点ID","文件名_日期","CheckCode=""" & Target_CheckCode & """" & TimeID_Range_String),"00000000-000000") _
		&"] " & Target_CheckCode _
	)

	'CheckCode 重新定位
	Target_CheckCode = Dfirst("CheckCode","文件名_日期","时点ID=" & Dmax("时点ID","文件名_日期","文件ID<900000"))

	Calling=Diary_Add("Message", _
		"CheckCode_Report: [" & _
		Format(Dmin("时点ID","文件名_日期","CheckCode=""" & Target_CheckCode & """" & TimeID_Range_String),"00000000-000000") _
		& ">>" & _
		Format(Dmax("时点ID","文件名_日期","CheckCode=""" & Target_CheckCode & """" & TimeID_Range_String),"00000000-000000") _
		&"] " & Target_CheckCode _
	)
End Function


