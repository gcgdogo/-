Option Explicit

Static CallDB_Application as Application
Static App_HasDiary as Boolean , CallDB_HasDiary as Boolean

Private Function Set_DBVersions_Direction() as String : Set_DBVersions_Direction = "F:\Database\数据库版本识别\DB_Versions.accdb" : End Function

Function CallDB_All_in_One(GroupName as String , MacroName as String) as Variant
End Function

Function CallDB_Connect(GroupName as String) as Variant
	
	'检测当前数据库是否有 Diary    On Error Goto App_Diary_Fail
	'检测目标数据库是否有 Diary    On Error Goto CallDB_Diary_Fail

	'检测 Diary 是否存在, 建立 Diary 连接

End Function

'Example:
'[Call]>>LangxinData: 获取文件版本 Get_FileVersion = LangxinData_1.12.1 [Hub Mod;Diary].accdb
'[Call]>>LangxinData: Diary_Application_Set : Diary 更新目标设置 / 未检测到 Diary_Application_Set
'[Call]>>LangxinData: <运行> 目标宏
'[Call:LangxinData] ~~~~~~~