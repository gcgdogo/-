'识别目标 ( ID , 添加时间 , 文件组 , 文件夹 , 文件名规则 , 最新发现时间 , 最新版本号 , 发现时间匹配 , 版本号匹配)
'识别结果 ( ID , 发现时间 , 文件组 , 文件夹 , 文件名 , 符合规则 , 版本号 , 存在)


Function Get_File_Names(Search_Origin_Table as String , TargetTable as String , Time_Search_Started as Variant)
    Dim ADO_rs as ADODB.Recordset
    Set ADO_rs = new ADODB.Recordset

    'Set:ADO_rs source=Search_Origin_Table

    'Do:ADO_rs=EOF
        call Search_Folder(ADO_rs!文件组 , ADO_rs!文件夹 , ADO_rs!文件名规则 , TargetTable)


End Function

Private Sub Search_File_Exist(TargetTable as String)
    Dim ADO_rs as ADODB.Recordset
    Set ADO_rs = new ADODB.Recordset

    'Set:ADO_rs source=TargetTable

    'Do:ADO_rs.EOF if ADO_rs!存在=true if dir()="" then ADO_rs!存在=false


End Sub

Private Sub Search_Folder(GroupName as String , FolderName as String , Str_Expression as String , TargetTable as String , Time_Search_Started as Variant)

    Dim FileNames(500) as String , FileNamesCount as integer , ThisFileName as String
    Dim ADO_rs as ADODB.Recordset
    Dim RE_PreEdit as RegExp , R_Match as Match

    Dim Str_Expression_Edited as String

    Set ADO_rs = new ADODB.Recordset
    Set RE_PreEdit as new RegExp

    'RE_PreEdit初始设置
    RE_PreEdit.IgnoreCase=True
    RE_PreEdit.Global=True

    '为<>外的特殊字符增加转义字符 "([\$\(\)\[\]\{\}\.\+\?])(?![^<]*>)"
    RE_PreEdit.Pattern="([\$\(\)\[\]\{\}\.\+\?])(?![^<]*>)"
    Str_Expression_Edited = RE_PreEdit.Replace(Str_Expression,"\$1")

    '<ver>替换为提取文本
    RE_PreEdit.Pattern="<ver>"
    Str_Expression_Edited = RE_PreEdit.Replace(Str_Expression_Edited,"(?:[^\.0-9]*?|[^\.]*?[^\.0-9]+?)([0-9]+)(?:\.([0-9]+))+.*?")

    '删除<>
    RE_PreEdit.Pattern="<(.*?)>"
    Str_Expression_Edited = RE_PreEdit.Replace(Str_Expression_Edited,"$1")


    Time_Search_Started = Now()

    'Set:ADO_rs source=TargetTable
    ADO_rs.ActiveConnection=CurrentProject.Connection
    ADO_rs.Source="SELECT 文件名 From " & TargetTable & " WHERE 文件组 = """ & GroupName & """ and 文件夹 = """ & FolderName & """ and 存在=True"
    ADO_rs.CursorType=adOpenDynamic
    ADO_rs.LockType=adLockOptimistic
    ADO_rs.Open

    'Do:ADO_rs.EOF FileNames()=ADO_rs!文件名
    FileNamesCount=0
    Do While ADO_rs.EOF=false
        FileNamesCount=FileNamesCount+1
        FileNames(FileNamesCount)=ADO_rs!文件名
    Loop

    ADO_rs.Close

    ADO_rs.ActiveConnection=CurrentProject.Connection
    ADO_rs.Source= TargetTable
    ADO_rs.CursorType=adOpenDynamic
    ADO_rs.LockType=adLockOptimistic
    ADO_rs.Open
    
        'ThisFileName=dir(FolderName)  Do:ThisFileName=dir

            'if RE_PreEdit.Test (Thisfilename , Str_Expression) then Set R_Matchs = RE_PreEdit.Execute

            'For:If:FileNames=Thisfilename ThisFileName=""

            'ADO_rs add(GroupName,FolderName,ThisFileName)
            
            'FileNames(+1)=ThisFileName


End sub

        
        'Need to EDIT:
        'Sub => function Boolean
        'RE_P rename
        
        