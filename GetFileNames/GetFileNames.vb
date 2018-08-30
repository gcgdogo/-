'识别目标 ( ID , 添加时间 , 文件组 , 文件夹 , 文件名规则 , 最新发现时间 , 最新版本号 , 发现时间匹配 , 版本号匹配)
'识别结果 ( ID , 发现时间 , 文件组 , 文件夹 , 文件名 , 符合规则 , 版本号 , 存在)


Function Get_File_Names(Search_Origin_Table as String , TargetTable as String , Time_Search_Started as Variant , Force_VersionChange as Boolean) as Boolean
    Dim Version_Changed as Boolean
    Dim ADO_rs as ADODB.Recordset
    Set ADO_rs = new ADODB.Recordset

    'Set:ADO_rs source=Search_Origin_Table

    Version_Changed=Force_VersionChange

    'Do:ADO_rs=EOF
        call Search_Folder(ADO_rs!文件组 , ADO_rs!文件夹 , ADO_rs!文件名规则 , TargetTable)


End Function

Private Function Search_File_Exist(TargetTable as String) as Boolean
    Dim Version_Changed as Boolean
    Dim ADO_rs as ADODB.Recordset
    Set ADO_rs = new ADODB.Recordset

    'Set:ADO_rs source=TargetTable

    'Do:ADO_rs.EOF if ADO_rs!存在=true if dir()="" then ADO_rs!存在=false


End Function

Private Function Search_Folder(GroupName as String , FolderName as String , Str_Expression as String , TargetTable as String , Time_Search_Started as Variant) as Boolean

    Dim FileNames(500) as String , FileNamesCount as integer , ThisFileName as String
    Dim ADO_rs as ADODB.Recordset
    Dim R_Exp as RegExp , R_Match as Match

    Dim Str_Expression_Edited as String

    Set ADO_rs = new ADODB.Recordset
    Set R_Exp as new RegExp

    'RE_PreEdit初始设置
    R_Exp.IgnoreCase=True
    R_Exp.Global=True

    '为<>外的特殊字符增加转义字符 "([\$\(\)\[\]\{\}\.\+\?])(?![^<]*>)"
    R_Exp.Pattern="([\$\(\)\[\]\{\}\.\+\?])(?![^<]*>)"
    Str_Expression_Edited = R_Exp.Replace(Str_Expression,"\$1")

    '<ver>替换为提取文本
    R_Exp.Pattern="<ver>"
    Str_Expression_Edited = R_Exp.Replace(Str_Expression_Edited,"(?:[^\.0-9]*?|[^\.]*?[^\.0-9]+?)([0-9]+)(?:\.([0-9]+))+.*?")

    '删除<>
    R_Exp.Pattern="<(.*?)>"
    Str_Expression_Edited = R_Exp.Replace(Str_Expression_Edited,"$1")


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

            'For:If:FileNames=Thisfilename ThisFileName=""

            'if R_Exp.Test (Thisfilename , Str_Expression) then Set R_Matchs = R_Exp.Execute

            'ADO_rs add(GroupName,FolderName,ThisFileName)
            
            'FileNames(+1)=ThisFileName


End sub      