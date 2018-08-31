'如果输入的 GroupName="" 跳过计算刷新过程,直接返回文件信息

Function Get_FileVersion(GroupName as String , Return_FullPath as Boolean) as String
    Dim Version_Changed as Boolean
    Dim TargetDataName as String
    Static Saved_GroupName as String

    if GroupName<>"" then
    	Saved_GroupName=GroupName
		Version_Changed = Upadate_FileVersions( "识别规则" , "识别文件列表" )
		if Version_Changed=True then Application.DoCmd.RunMacro("识别文件汇总")
	End If

	If Return_FullPath=True Then TargetDataName="匹配完整路径"
	If Return_FullPath=False Then TargetDataName="匹配文件名"

	Get_FileVersion=DFirst(TargetDataName,"识别规则","文件组=""" & Saved_GroupName & """")
    LastSearch_Timer=Timer()
End Function