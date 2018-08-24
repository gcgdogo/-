Sub Get_File_Names(GroupName as string , FolderName as string , NameStyle as string , TargetTable as string)

  Dim FileNames(500) as string , FileNamesCount as integer , ThisFileName as string
  Dim ADO_rs as adodb.recordset
  Dim Chk_String as RegExp
  Set ADO_rs = new adodb.recordset
  set Chk_String as new RegExp

  'Set:ADO_rs source=TargetTable

  'Do:ADO_rs.EOF FileNames()=ADO_rs!FileName

  'ThisFileName=dir(FolderName)  Do:ThisFileName=dir

    'Chk_String.Test (Thisfilename , NameStyle)

    'For:If:FileNames=Thisfilename ThisFileName=""

    'ADO_rs add(GroupName,FolderName,ThisFileName)

    'FileName=ThisFileName

    'testajldfjasdlkfjalkfdj
    


End sub

