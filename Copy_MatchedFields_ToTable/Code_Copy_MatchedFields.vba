'此模块用于快速追加数据, 能自动选择匹配的字段, 防止两表字段不同导致需要在 Select 后填写大量名称
'暂时用于工资数据库中整合数据

'以后再不会这样了:
'             select
'                 "员工工资套" as 数据来源,
'                 年月,
'                 应发工资 as 系统应发,
'                 次数,人员类别,员工编号,姓名,身份证号码,部门名称,部门,参考部门
'                     ,岗位类别,岗位级别,个人级别,职级,档位,薪酬级别,岗位级别工资,绩效工资
'                     ,加班工资,特殊奖罚,优秀员工补贴,大学生补贴,车补,病事假,岗位能力工资
'                     ,司龄工资,临时补贴,技术津贴,项目奖,等级工资,应税工资,个人所得税,实发工资
'                     ,会费,养老保险,公积金,医疗保险,医疗保险调整额,失业保险,公积金调整额,其他
'                     ,质量扣款,行政扣款,餐补,保健,大额医保,未发零头,职务（工种）,一级类别
'                     ,二级类别,三级类别,岗位体系,岗位域,岗位族,岗位名称,岗位名称1,业务群机构
'                     ,单位级机构,部级机构,室级机构,工段级机构,发放分类
'             from 员工工资套

Option Compare Database
Option Explicit

Function Copy_MatchedFields(SourceTable as String , TargetTable as String , Optional ForeAddon_FieldList as String = "")
    Dim Matched_FieldList as String , Combined_FieldList as String
    Dim ADO_rs As New ADODB.Recordset
    Dim FieldNames_Dict As New Dictionary
    
    Dim I as Integer

    'ForeAddon_FieldList 末尾逗号判定
    If ForeAddon_FieldList<>"" Then
        If Right(ForeAddon_FieldList,1)<>"," Then ForeAddon_FieldList = ForeAddon_FieldList & ","
    End If


    Matched_FieldList = ""

    '遍历 SourceTable 字段
    ADO_rs.ActiveConnection = CurrentProject.Connection
    ADO_rs.Source = SourceTable
    ADO_rs.CursorType = adOpenStatic
    ADO_rs.LockType = adLockReadOnly
    ADO_rs.Open

    For I = 0 To ADO_rs.Fields.Count - 1
        If FieldNames_Dict.Exists(ADO_rs.Fields(I).Name) Then
            '在字典中标注已匹配项目, 便于Debug
            FieldNames_Dict(ADO_rs.Fields(I).Name) = FieldNames_Dict(ADO_rs.Fields(I).Name) & " : " & ADO_rs.Source & "_" & I
            Matched_FieldList = Matched_FieldList & "[" & ADO_rs.Fields(I).Name & "],"
        Else
            FieldNames_Dict.Add ADO_rs.Fields(I).Name , ADO_rs.Source & "_" & I
        End If
    Next
    ADO_rs.Close

    '遍历 TargetTable 字段
    ADO_rs.ActiveConnection = CurrentProject.Connection
    ADO_rs.Source = TargetTable
    ADO_rs.CursorType = adOpenStatic
    ADO_rs.LockType = adLockReadOnly
    ADO_rs.Open

    For I = 0 To ADO_rs.Fields.Count - 1
        If FieldNames_Dict.Exists(ADO_rs.Fields(I).Name) Then
            '在字典中标注已匹配项目, 便于Debug
            FieldNames_Dict(ADO_rs.Fields(I).Name) = FieldNames_Dict(ADO_rs.Fields(I).Name) & " : " & ADO_rs.Source & "_" & I
            Matched_FieldList = Matched_FieldList & "[" & ADO_rs.Fields(I).Name & "],"
        Else
            FieldNames_Dict.Add ADO_rs.Fields(I).Name , ADO_rs.Source & "_" & I
        End If
    Next
    ADO_rs.Close

    Combined_FieldList = ForeAddon_FieldList & Matched_FieldList

    '字段列表判定
    If Combined_FieldList = "" Then
        Msgbox("未找到需导入字段: " & vbCrLf & SourceTable & vbCrLf & TargetTable)
        Exit Function
    End If

    '末尾逗号处理
    If Right(Combined_FieldList,1) = "," Then Combined_FieldList = Left(Combined_FieldList , Len(Combined_FieldList) - 1)

    '执行查询
    ADO_rs.ActiveConnection = CurrentProject.Connection
    ADO_rs.Source = "INSERT INTO " & TargetTable & vbCrLf & _
                    "SELECT " & Combined_FieldList & vbCrLf & _
                    "FROM " & SourceTable
    ADO_rs.CursorType = adOpenDynamic
    ADO_rs.LockType = adLockOptimistic
    ADO_rs.Open

End Function