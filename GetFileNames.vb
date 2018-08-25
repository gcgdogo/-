Function Get_File_Names(Search_Origin as string , TargetTable as string)
    Dim ADO_rs as adodb.recordset
    Set ADO_rs = new adodb.recordset

    'Set:ADO_rs source=Search_Origin

    'Do:ADO_rs=EOF
        call Search_Folder(ADO_rs!版本组 , ADO_rs!文件夹 , ADO_rs!文件名规则 , TargetTable)


End Function


Private Sub Search_Folder(GroupName as string , FolderName as string , NameStyle as string , TargetTable as string)

    Dim FileNames(500) as string , FileNamesCount as integer , ThisFileName as string
    Dim ADO_rs as adodb.recordset
    Dim Chk_String as RegExp
    Set ADO_rs = new adodb.recordset
    set Chk_String as new RegExp

    'Set:ADO_rs source=TargetTable

    'Do:ADO_rs.EOF FileNames()=ADO_rs!文件名

        'ThisFileName=dir(FolderName)  Do:ThisFileName=dir

        'Chk_String.Test (Thisfilename , NameStyle)

        'For:If:FileNames=Thisfilename ThisFileName=""

        'ADO_rs add(GroupName,FolderName,ThisFileName)

        'FileName=ThisFileName

End sub
