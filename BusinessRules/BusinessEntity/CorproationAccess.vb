Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class corporationAccess
    'Implements IDisposable


    Public Const Table_corporation As String = "corporation"
    Public Const Table_consultation As String = "consultation"


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Corporation As SqlDataAdapter
    Private dsCommand_Consultation As SqlDataAdapter

    Private GetcorporationInfoCommand As SqlCommand
    Private GetConsultationInfoCommand As SqlCommand

    Private GetCorporationMaxCodeCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction



    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        ' 实例化适配器
        dsCommand_Corporation = New SqlDataAdapter()
        dsCommand_Consultation = New SqlDataAdapter()


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetcorporationInfo("null", "null")
    End Sub


    '获取申请担保最大企业ID
    Public Function GetCorporationMaxCode() As String
        Dim MaxCode As String
        If GetCorporationMaxCodeCommand Is Nothing Then

            GetCorporationMaxCodeCommand = New SqlCommand("GetCorporationMaxCode", conn)
            GetCorporationMaxCodeCommand.CommandType = CommandType.StoredProcedure

            GetCorporationMaxCodeCommand.Parameters.Add(New SqlParameter("@CorporationMaxCode", SqlDbType.BigInt))
            GetCorporationMaxCodeCommand.Parameters.Item("@CorporationMaxCode").Direction = ParameterDirection.Output
        End If

        GetCorporationMaxCodeCommand.Transaction = ts
        GetCorporationMaxCodeCommand.ExecuteNonQuery()


        MaxCode = CStr(GetCorporationMaxCodeCommand.Parameters.Item("@CorporationMaxCode").Value)
        Dim tmpLen As Integer
        Dim i As Integer
        tmpLen = 5 - Len(MaxCode)
        For i = 0 To tmpLen - 1
            MaxCode = "0" & MaxCode
        Next

        GetCorporationMaxCode = MaxCode
    End Function

    '获取保证企业最大企业ID
    Public Function GetCorporationMaxCode_Guarantee() As String
        Dim MaxCode As String
        If GetCorporationMaxCodeCommand Is Nothing Then

            GetCorporationMaxCodeCommand = New SqlCommand("GetCorporationMaxCode_guarantee", conn)
            GetCorporationMaxCodeCommand.CommandType = CommandType.StoredProcedure

            GetCorporationMaxCodeCommand.Parameters.Add(New SqlParameter("@CorporationMaxCode", SqlDbType.BigInt))
            GetCorporationMaxCodeCommand.Parameters.Item("@CorporationMaxCode").Direction = ParameterDirection.Output
        End If

        GetCorporationMaxCodeCommand.Transaction = ts
        GetCorporationMaxCodeCommand.ExecuteNonQuery()


        MaxCode = CStr(GetCorporationMaxCodeCommand.Parameters.Item("@CorporationMaxCode").Value)
        Dim tmpLen As Integer
        Dim i As Integer
        tmpLen = 5 - Len(MaxCode)
        For i = 0 To tmpLen - 1
            MaxCode = "0" & MaxCode
        Next
        MaxCode = "A" & MaxCode.Substring(1, 4)
        GetCorporationMaxCode_Guarantee = MaxCode
    End Function

    '获取该企业的项目编码
    Public Function GetProjectCode(ByVal corporationCode As String) As String
        Dim tmpYear As Integer = Year(Now)
        Dim Project As New Project(conn, ts)
        'Dim strSql As String = "{corporation_code=" & "'" & corporationCode & "'" & " and year(create_date)=" & tmpYear & " order by project_code}"

        '根据受理时间的年份和项目编码的第6，7为决定项目编码，而不根据APPLY_DATE
        Dim strSql As String = "{corporation_code=" & "'" & corporationCode & "'" & " and substring(project_code,6,2)=" & Mid(CStr(tmpYear), 3, 2) & " order by project_code}"
        Dim dsTempProject As DataSet = Project.GetProjectInfo(strSql)

        '获取该企业最大项目编码
        Dim tmpMaxProjectNum As Integer
        Dim rowNum As Integer
        Dim tmpProjectNum As String

        If dsTempProject.Tables(0).Rows.Count = 0 Then

            '如果该企业第一次建立项目,项目编码从1开始
            tmpProjectNum = "1"

        Else

            rowNum = dsTempProject.Tables(0).Rows.Count - 1
            tmpMaxProjectNum = CInt(Mid(dsTempProject.Tables(0).Rows(rowNum).Item("project_code"), 8, 2))

            '项目编码为最大项目编码+1
            tmpProjectNum = CStr(tmpMaxProjectNum + 1)
        End If

        If Len(tmpProjectNum) = 1 Then
            tmpProjectNum = "0" & tmpProjectNum
        End If
        Dim tmpCorporationCode As String = StrCorporationCode(corporationCode)
        GetProjectCode = tmpCorporationCode & Mid(CStr(tmpYear), 3, 2) & tmpProjectNum

    End Function

    Private Function StrCorporationCode(ByVal corporationCode As String) As String
        Dim tmpCorporationCode As String = corporationCode
        Dim tmpLen As Integer = 5 - Len(tmpCorporationCode)
        Dim i As Integer
        For i = 0 To tmpLen - 1
            tmpCorporationCode = "0" & tmpCorporationCode
        Next
        Return tmpCorporationCode
    End Function

    '获取企业信息
    Public Function GetcorporationInfo(ByVal strSQL_Condition_Corporation As String, ByVal strSQL_Condition_Consultation As String) As DataSet

        Dim tempDs As New DataSet()

        If GetcorporationInfoCommand Is Nothing Then

            GetcorporationInfoCommand = New SqlCommand("GetcorporationInfo", conn)
            GetcorporationInfoCommand.CommandType = CommandType.StoredProcedure
            GetcorporationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Corporation
            .SelectCommand = GetcorporationInfoCommand
            .SelectCommand.Transaction = ts
            GetcorporationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Corporation
            .Fill(tempDs, Table_corporation)
        End With

        If GetConsultationInfoCommand Is Nothing Then

            GetConsultationInfoCommand = New SqlCommand("GetConsultationInfo", conn)
            GetConsultationInfoCommand.CommandType = CommandType.StoredProcedure
            GetConsultationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Consultation
            .SelectCommand = GetConsultationInfoCommand
            .SelectCommand.Transaction = ts
            GetConsultationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Consultation

            .Fill(tempDs, Table_consultation)
        End With
        'tempDs.Tables(1).Columns.Add("test", GetType(String))
        'Dim rl As DataRelation
        'rl = New DataRelation("CorporationConsultation", tempDs.Tables(0).Columns("corporation_code"), tempDs.Tables(1).Columns("corporation_code"), False)
        'tempDs.Relations.Add(rl)
        'tempDs.Tables(1).Columns.Add("corporation_name_corporation", GetType(String), "Parent(CorporationConsultation).corporation_name")
        GetcorporationInfo = tempDs



    End Function

    '更新企业信息
    Public Function UpdateCorporation(ByVal corporationSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Corporation)

        With dsCommand_Corporation
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(corporationSet, Table_corporation)
        End With

    End Function



    '更新企业咨询信息
    Public Function UpdateConsultation(ByVal ConsultationSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Consultation)

        With dsCommand_Consultation
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ConsultationSet, Table_consultation)

        End With

    End Function

    '更新企业及咨询信息
    Public Function UpdateCorCon(ByVal ConsultationSet As DataSet)

        If ConsultationSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If ConsultationSet.HasChanges = False Then
            Exit Function
        End If

        '删除操作
        If IsNothing(ConsultationSet.GetChanges(DataRowState.Deleted)) = False Then
            '先删明细表，再删主表
            UpdateConsultation(ConsultationSet.GetChanges(DataRowState.Deleted))
            UpdateCorporation(ConsultationSet.GetChanges(DataRowState.Deleted))

        End If

        '新增操作
        If IsNothing(ConsultationSet.GetChanges(DataRowState.Added)) = False Then

            Dim i, j As Integer

            '获取企业最大代码
            Dim MaxCode As String = GetCorporationMaxCode()
            Dim tmpLen As Integer

            '定义临时适配器（供填充企业信息表）
            Dim dsCommand_Temp As New SqlDataAdapter()

            '定义临时查询命令（供临时适配器用）
            Dim GetCorporationByNameCommand As SqlCommand
            GetCorporationByNameCommand = New SqlCommand("GetcorporationInfo", conn)
            GetCorporationByNameCommand.CommandType = CommandType.StoredProcedure
            GetCorporationByNameCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

            GetCorporationByNameCommand.Transaction = ts


            Dim dsCorp As DataSet

            '批量新增企业信息采集（客户端输入的企业信息采集为多条）
            Dim tempRowConsultation As DataRow
            For i = 0 To ConsultationSet.Tables(1).Rows.Count - 1
                tempRowConsultation = ConsultationSet.Tables(1).Rows(i)
                If tempRowConsultation.RowState = DataRowState.Added Then
                    Dim tmpCorporationName As String = Trim(tempRowConsultation.Item("corporation_name"))
                    Dim strSql As String = "{corporation_name=" & "'" & tmpCorporationName & "'" & " and corporation_type='1'}"

                    Dim newrow As DataRow

                    '获取与企业名相匹配的企业信息
                    dsCorp = GetcorporationInfo(strSql, "NULL")

                    '如果新增的咨询企业是新的企业（企业信息记录数为零）
                    If dsCorp.Tables(0).Rows.Count = 0 Then

                        With tempRowConsultation
                            '新增企业信息()
                            newrow = ConsultationSet.Tables(0).NewRow()

                            newrow.Item("corporation_code") = MaxCode
                            newrow.Item("corporation_type") = .Item("corporation_type")
                            newrow.Item("corporation_name") = .Item("corporation_name")
                            newrow.Item("register_address") = .Item("register_address")
                            newrow.Item("district_name") = .Item("district_name")
                            newrow.Item("district_parent_name") = .Item("district_parent_name")
                            newrow.Item("proprietorship_type") = .Item("proprietorship_type")
                            newrow.Item("industry_type") = .Item("industry_type")
                            newrow.Item("contact_person") = .Item("contact_person")
                            newrow.Item("job") = .Item("job")
                            newrow.Item("phone_num") = .Item("phone_num")
                            newrow.Item("fax") = .Item("fax")
                            newrow.Item("web_site") = .Item("web_site")
                            newrow.Item("e_mail") = .Item("e_mail")
                            newrow.Item("create_person") = .Item("create_person")
                            newrow.Item("create_date") = .Item("create_date")
                            newrow.Item("found_date") = .Item("found_date")
                            newrow.Item("technology_type") = .Item("technology_type")
                            newrow.Item("mobile") = .Item("mobile")

                            newrow.Item("credit_grade") = .Item("credit_grade") ' qxd add 2004-2-6
                            newrow.Item("corporation_type_ex") = .Item("corporation_type_ex") ' yjf add 2009-9-10
                            newrow.Item("if_import") = .Item("if_import") 'yansm add 2013-11-25
                            newrow.Item("if_technology") = .Item("if_technology") 'yansm add 2013-11-25
                            newrow.Item("indus_type") = .Item("indus_type") 'yansm add 2013-11-25
                            'newrow.Item("referenciae_type") = .Item("referencia_type") 'yansm add 2013-11-29



                            ConsultationSet.Tables(0).Rows.Add(newrow)

                            '找到该企业在咨询表中新增的那条记录

                            '填充企业编码和咨询次数

                            .Item("corporation_code") = MaxCode
                            .Item("serial_num") = 1

                        End With

                        UpdateCorporation(ConsultationSet.GetChanges(DataRowState.Added))
                        UpdateConsultation(ConsultationSet.GetChanges(DataRowState.Added))
                    Else
                        '如果新增的企业存在
                        '修改该企业信息

                        Dim dsTempCorporation As New DataSet()

                        '填充企业信息表
                        With dsCommand_Temp
                            .SelectCommand = GetCorporationByNameCommand
                            GetCorporationByNameCommand.Parameters("@Condition").Value = strSql
                            .Fill(dsTempCorporation, Table_corporation)
                        End With

                        Dim corporationCode As String = dsTempCorporation.Tables(0).Rows(0).Item("corporation_code")

                        With tempRowConsultation

                            '修改企业信息表中的某些字段的值
                            newrow = dsTempCorporation.Tables(0).Rows(0)

                            newrow.Item("corporation_type") = .Item("corporation_type")
                            newrow.Item("corporation_name") = .Item("corporation_name")
                            newrow.Item("register_address") = .Item("register_address")
                            newrow.Item("district_name") = .Item("district_name")
                            newrow.Item("district_parent_name") = .Item("district_parent_name")
                            newrow.Item("proprietorship_type") = .Item("proprietorship_type")
                            newrow.Item("industry_type") = .Item("industry_type")
                            newrow.Item("contact_person") = .Item("contact_person")
                            newrow.Item("job") = .Item("job")
                            newrow.Item("phone_num") = .Item("phone_num")
                            newrow.Item("fax") = .Item("fax")
                            newrow.Item("web_site") = .Item("web_site")
                            newrow.Item("e_mail") = .Item("e_mail")
                            newrow.Item("create_person") = .Item("create_person")
                            newrow.Item("create_date") = .Item("create_date")
                            newrow.Item("found_date") = .Item("found_date")
                            newrow.Item("technology_type") = .Item("technology_type")
                            newrow.Item("mobile") = .Item("mobile")
                            newrow.Item("corporation_type_ex") = .Item("corporation_type_ex") ' yjf add 2009-9-10
                            newrow.Item("if_import") = .Item("if_import") 'yansm add 2013-11-25
                            newrow.Item("if_technology") = .Item("if_technology") 'yansm add 2013-11-25
                            newrow.Item("indus_type") = .Item("indus_type") 'yansm add 2013-11-25
                            'newrow.Item("referencia_type") = .Item("referencia_type") 'yansm add 2013-11-29
                        End With

                        '按咨询次数排序，查出该企业咨询信息
                        strSql = "{corporation_code=" & "'" & corporationCode & "'" & " order by serial_num" & "}"

                        Dim RowNum, MaxSerialNum As Integer

                        '取出最大咨询次数
                        Dim dsTemp As DataSet = GetcorporationInfo("null", strSql)
                        RowNum = dsTemp.Tables(1).Rows.Count
                        MaxSerialNum = dsTemp.Tables(1).Rows(RowNum - 1).Item("serial_num") + 1


                        '填充企业编码和咨询次数
                        With tempRowConsultation
                            .Item("corporation_code") = corporationCode
                            .Item("serial_num") = MaxSerialNum
                        End With


                        UpdateCorporation(dsTempCorporation)
                        UpdateConsultation(ConsultationSet.GetChanges(DataRowState.Added))

                    End If

                    MaxCode = CStr(CInt(MaxCode) + 1)
                    tmpLen = 5 - Len(MaxCode)
                    For j = 0 To tmpLen - 1
                        MaxCode = "0" & MaxCode
                    Next

                End If


            Next


        End If

        '修改操作
        If IsNothing(ConsultationSet.GetChanges(DataRowState.Modified)) = False Then

            If Not ConsultationSet.Tables(1).GetChanges(DataRowState.Modified) Is Nothing Then
                Dim i As Integer
                Dim strSql As String
                Dim tempRowConsultation, newRow As DataRow
                Dim tmpCorporationCode As String

                '定义临时适配器（供填充企业信息表）
                Dim dsCommand_Temp As New SqlDataAdapter()

                '定义临时查询命令（供临时适配器用）
                Dim GetCorporationByNameCommand As SqlCommand
                GetCorporationByNameCommand = New SqlCommand("GetcorporationInfo", conn)
                GetCorporationByNameCommand.CommandType = CommandType.StoredProcedure
                GetCorporationByNameCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

                GetCorporationByNameCommand.Transaction = ts

                For i = 0 To ConsultationSet.GetChanges(DataRowState.Modified).Tables(1).Rows.Count - 1
                    Dim dsTempCorporation As New DataSet()
                    tmpCorporationCode = ConsultationSet.GetChanges(DataRowState.Modified).Tables(1).Rows(i).Item("corporation_code")
                    tempRowConsultation = ConsultationSet.GetChanges(DataRowState.Modified).Tables(1).Rows(i)

                    strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"

                    '填充企业信息表
                    With dsCommand_Temp
                        .SelectCommand = GetCorporationByNameCommand
                        GetCorporationByNameCommand.Parameters("@Condition").Value = strSql
                        .Fill(dsTempCorporation, Table_corporation)
                    End With

                    With tempRowConsultation

                        '修改企业信息表中的某些字段的值
                        newrow = dsTempCorporation.Tables(0).Rows(0)

                        newrow.Item("corporation_type") = .Item("corporation_type")
                        newrow.Item("corporation_name") = .Item("corporation_name")
                        newrow.Item("register_address") = .Item("register_address")
                        newrow.Item("district_name") = .Item("district_name")
                        newrow.Item("district_parent_name") = .Item("district_parent_name")
                        newrow.Item("proprietorship_type") = .Item("proprietorship_type")
                        newrow.Item("industry_type") = .Item("industry_type")
                        newrow.Item("contact_person") = .Item("contact_person")
                        newrow.Item("job") = .Item("job")
                        newrow.Item("phone_num") = .Item("phone_num")
                        newrow.Item("fax") = .Item("fax")
                        newrow.Item("web_site") = .Item("web_site")
                        newrow.Item("e_mail") = .Item("e_mail")
                        newrow.Item("create_person") = .Item("create_person")
                        newrow.Item("create_date") = .Item("create_date")
                        newrow.Item("found_date") = .Item("found_date")
                        newrow.Item("technology_type") = .Item("technology_type")
                        newrow.Item("mobile") = .Item("mobile")
                        newRow.Item("corporation_type_ex") = .Item("corporation_type_ex") ' yjf add 2009-9-10
                        newRow.Item("if_import") = .Item("if_import") 'yansm add 2013-11-25
                        newRow.Item("if_technology") = .Item("if_technology") 'yansm add 2013-11-25
                        newRow.Item("indus_type") = .Item("indus_type") 'yansm add 2013-11-25
                        'newRow.Item("referencia_type") = .Item("referencia_type")  'yansm add 2013-11-29


                    End With

                    UpdateCorporation(dsTempCorporation)
                Next


                UpdateConsultation(ConsultationSet.GetChanges(DataRowState.Modified))

            Else

                UpdateCorporation(ConsultationSet.GetChanges(DataRowState.Modified))
                UpdateConsultation(ConsultationSet.GetChanges(DataRowState.Modified))
                ConsultationSet.AcceptChanges()
            End If
        End If

    End Function
End Class