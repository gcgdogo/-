'Search_Origin_Table = 识别规则 ( ID , 添加时间 , 运行 , 文件组 , 文件夹 , 文件名规则 , 文件夹内现存文件 , 组内最新时间 , 组内版本号 , 组内匹配数量 , 匹配文件名 , 匹配完整路径)
'TargetTable = 识别文件列表 ( ID , 发现时间 , 文件组 , 文件夹 , 文件名 , 版本号 , 存在)

Function Upadate_FileVersions(Search_Origin_Table as String , TargetTable as String ) as Boolean
    Dim Version_Changed as Boolean
    Dim ADO_rs as New ADODB.Recordset

    Version_Changed=False

    '调用 Search_File_Exist
    Version_Changed=Version_Changed or Search_File_Exist(TargetTable)

    'Set:ADO_rs source=Search_Origin_Table
    ADO_rs.ActiveConnection=CurrentProject.Connection
    ADO_rs.Source="SELECT * FROM " & Search_Origin_Table & " WHERE 运行=True"
    ADO_rs.CursorType=adOpenStatic
    ADO_rs.LockType=adLockReadOnly
    ADO_rs.Open
    'Do:ADO_rs=EOF
    Do While ADO_rs.EOF=False
        'Private Function Search_Folder(GroupName as String , FolderName as String , Str_Expression as String , TargetTable as String ) as Boolean
        Version_Changed=Version_Changed or Search_Folder( ADO_rs!文件组 , ADO_rs!文件夹 , ADO_rs!文件名规则 , TargetTable)
        ADO_rs.MoveNext
    Loop

    Upadate_FileVersions=Version_Changed
End Function

Private Function Search_File_Exist(TargetTable as String) as Boolean
    Dim Version_Changed as Boolean
    Dim ADO_rs as New ADODB.Recordset
    Dim Dir_String as String

    Version_Changed=False

    'Set:ADO_rs source=TargetTable
    ADO_rs.ActiveConnection=CurrentProject.Connection
    ADO_rs.Source=TargetTable
    ADO_rs.CursorType=adOpenDynamic
    ADO_rs.LockType=adLockOptimistic
    ADO_rs.Open
    'Do:ADO_rs.EOF if ADO_rs!存在=true if Dir()="" Then ADO_rs!存在=False
    Do While ADO_rs.EOF=False
        if ADO_rs!存在=True Then
            if right(ADO_rs!文件夹,1)="\" then Dir_String= ADO_rs!文件夹 & ADO_rs!文件名 else Dir_String= ADO_rs!文件夹 & "\" & ADO_rs!文件名
            if Dir(Dir_String)="" then
                Version_Changed=True
                ADO_rs!存在=False
                ADO_rs.Update
            end if
        end if
        ADO_rs.MoveNext
    Loop 

    ADO_rs.Close

    Search_File_Exist=Version_Changed
End Function

Private Function Search_Folder(GroupName as String , FolderName as String , Str_Expression as String , TargetTable as String ) as Boolean

    Dim FileNames(500) as String , FileNamesCount as Integer , FileNamesFinger as Integer , ThisFileName as String
    Dim I as Integer

    Dim ADO_rs as New ADODB.Recordset
    Dim R_Exp as New RegExp , R_SubMatchs as Match.SubMatches

    Dim Str_Expression_Edited as String , VerPhaseString as String , Time_Search_Started as Variant
    Dim Version_Changed as Boolean

    'RE_PreEdit初始设置
    R_Exp.IgnoreCase=True
    R_Exp.Global=True

    '为<>外的特殊字符增加转义字符 "([\$\(\)\[\]\{\}\.\+\?])(?![^<]*>)"
    R_Exp.Pattern="([\$\(\)\[\]\{\}\.\+\?])(?![^<]*>)"
    Str_Expression_Edited = R_Exp.Replace(Str_Expression,"\$1")

    '<ver>替换为提取文本 , 共支持8个(.) 9串数字
    R_Exp.Pattern="<ver>"
    Str_Expression_Edited = R_Exp.Replace(Str_Expression_Edited,"(?:[^\.0-9]*?|[^\.]*?[^\.0-9]+?)([0-9]+)(?:\.([0-9]+))?(?:\.([0-9]+))?(?:\.([0-9]+))?(?:\.([0-9]+))?(?:\.([0-9]+))?(?:\.([0-9]+))?(?:\.([0-9]+))?(?:\.([0-9]+))?.*?")

    '删除<>
    R_Exp.Pattern="<(.*?)>"
    Str_Expression_Edited = R_Exp.Replace(Str_Expression_Edited,"$1")

    '增加首尾限定符
    Str_Expression_Edited = "^" & Str_Expression_Edited & "$"

    'Set:ADO_rs source=TargetTable
    ADO_rs.ActiveConnection=CurrentProject.Connection
    ADO_rs.Source="SELECT 文件名 From " & TargetTable & " WHERE 文件组 = """ & GroupName & """ and 文件夹 = """ & FolderName & """ and 存在=True"
    ADO_rs.CursorType=adOpenDynamic
    ADO_rs.LockType=adLockOptimistic
    ADO_rs.Open

    'Do:ADO_rs.EOF FileNames()=ADO_rs!文件名
    FileNamesCount=0
    Do While ADO_rs.EOF=False
        FileNamesCount=FileNamesCount+1
        FileNames(FileNamesCount)=ADO_rs!文件名
        ADO_rs.MoveNext
    Loop

    ADO_rs.Close
    
    '文件夹循环开始前变量赋值
    Time_Search_Started = Now()
    Version_Changed=False

    '正则表达式条件设置
    R_Exp.Pattern=Str_Expression_Edited

    '重新打开ADO用来写入新发现文件
    ADO_rs.ActiveConnection=CurrentProject.Connection
    ADO_rs.Source= TargetTable
    ADO_rs.CursorType=adOpenDynamic
    ADO_rs.LockType=adLockOptimistic
    ADO_rs.Open

    'ThisFileName=Dir(FolderName)  Do:ThisFileName=Dir
    ThisFileName=Dir(FolderName)
    Do While ThisFileName<>""
        '查找历史项目 For:If:FileNames=Thisfilename ThisFileName=""
        For FileNamesFinger=1 To FileNamesCount
            If FileNames(FileNamesFinger)=ThisFileName Then Exit For
        Next
        '无匹配项 且符合R_Exp 则新增
        If FileNamesFinger>FileNamesCount and R_Exp.Test(ThisFileName)=True Then
            if FileNamesFinger<>FileNamesCount+1 then
                Msgbox("ERROR::FileNamesFinger 循环出现错误")
                Exit Function
            end if

            '识别结果 ( ID , 发现时间 , 文件组 , 文件夹 , 文件名 , 版本号 , 存在)
            ADO_rs.AddNew

            ADO_rs!发现时间=Time_Search_Started
            ADO_rs!文件组=GroupName
            ADO_rs!文件夹=FolderName
            ADO_rs!文件名=ThisFileName
            ADO_rs!存在=True

            Version_Changed=True

            'if R_Exp.Test (Thisfilename , Str_Expression) Then 
            'Set R_SubMatchs = R_Exp.Execute.item0.SubMatches

                'RegExp系列的 Item 都是从零开始计数的 即范围为[ 0 ~ Count-1 ]
                Set R_SubMatchs = R_Exp.Execute(ThisFileName).Item(0).SubMatches

                For I=0 To R_SubMatchs.Count-1 

                    '空值跳过,即不添加00000000
                    if R_SubMatchs(I)<>"" then
                        '版本号小节统一成8个数字
                        VerPhaseString=Right( "00000000" & R_SubMatchs(I) , 8 )
                        if I=0 then ADO_rs!版本号=VerPhaseString else ADO_rs!版本号=ADO_rs!版本号 & "." & VerPhaseString
                    end if
                Next
            'ADO_rs add(GroupName,FolderName,ThisFileName)
            ADO_rs.Update
            'FileNames(+1)=ThisFileName
            FileNamesCount=FileNamesCount+1
            FileNames(FileNamesCount)=ThisFileName
        End If
        ThisFileName=Dir
    Loop
    ADO_rs.Close

    Search_Folder=Version_Changed
End Function
