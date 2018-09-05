Option Explicit

Static CallDB_Application as Application
Static App_HasDiary as Boolean , CallDB_HasDiary as Boolean

Private Function Set_DBVersions_Direction() as String : Set_DBVersions_Direction = "F:\Database\数据库版本识别\DB_Versions.accdb" : End Function

Function CallDB_All_in_One(GroupName as String , MacroName as String) as Variant
	Dim Calling as Variant
	Calling = CallDB_Connect(GroupName)
	CallDB_Application.DoCmd.RunMacro(MacroName)
	CallDB_Application.Quit
End Function

Function CallDB_Connect(GroupName as String) as Variant
	Dim TarDB_Direction as String , TarDB_Filename as String
	Dim Calling as Variant

	'初始值设定
	On Error Goto CallDB_FirstSet
		CallDB_Application.Quit
CallDB_FirstSet:
	On Error Goto 0
	Set CallDB_Application = New Access.Application

	'调用DBVersions,获取目标文件
	CallDB_Application.OpenCurrentDatabase(Set_DBVersions_Direction())
	TarDB_Direction = CallDB_Application.Run( "Get_FileVersion" , GroupName , True)
	TarDB_Filename = CallDB_Application.Run( "Get_FileVersion" , "" , False)

	'检测当前数据库是否有 Diary    On Error Goto App_Diary_Fail
	App_HasDiary = False
	CallDB_HasDiary = False
	On Error Goto App_Diary_Fail
		'确保本地 DairyAdd 调用过,避免 Diary_Application 没有初始化
		Calling = Diary_Add("Message", "[Call]>>" & GroupName & ": 获取文件版本 Get_FileVersion = " & TarDB_Filename)
		App_HasDiary = True
		'检测目标数据库是否有 Diary   建立 Diary 连接  On Error Goto CallDB_Diary_Fail
		On Error Goto CallDB_Diary_Fail
			Calling = CallDB_Application.Run("Diary_Application_Set" , Diary_Application)
			Calling = CallDB_Application.Run("Diary_HeadString_Set" , Diary_HeadString & "[Call:" & GroupName & "] ")
			CallDB_HasDiary = True
CallDB_Diary_Fail:
		On Error Goto App_Diary_Fail
		If CallDB_HasDiary = True Then
			Calling = Diary_Add("Message", "[Call]>>" & GroupName & ": Diary_Application_Set : Diary 更新目标设置")
		Else
			Calling = Diary_Add("Message", "[Call]>>" & GroupName & ": 未检测到 Diary_Application_Set")
		End If
	On Error Goto 0
End Function

'Example:
'[Call]>>LangxinData: 获取文件版本 Get_FileVersion = LangxinData_1.12.1 [Hub Mod;Diary].accdb
'[Call]>>LangxinData: Diary_Application_Set : Diary 更新目标设置 / 未检测到 Diary_Application_Set
'[Call]>>LangxinData: <运行> 目标宏
'[Call:LangxinData] ~~~~~~~