Sub Get_File_Names(GroupName as string , FolderName as string , NameStyle as string , TargetTable as string)

  	Dim FileNames(500) as string , FileNamesCount as integer , ThisFileName as string
  	Dim ADO_rs as adodb.recordset
  	Set ADO_rs = new adodb.recordset

    'Set:ADO_rs source=TargetTable

    'Do:ADO_rs.EOF FileNames()=ADO_rs!FileName

    'Do:ThisFileName=dir(FolderName)

        'Check:RegExp (Thisfilename , NameStyle)

 		'For:If:FileNames=Thisfilename ThisFileName=""

      	'ADO_rs add

      	'FileName=ThisFileName

      '123156
    


End sub
