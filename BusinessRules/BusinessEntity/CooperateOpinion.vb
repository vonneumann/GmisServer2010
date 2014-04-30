Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CooperateOpinion

    Public Const Table_Cooperate_Organization As String = "cooperate_organization"
    Public Const Table_Cooperate_Organization_Opinion As String = "cooperate_organization_opinion"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_CooperateOrganization As SqlDataAdapter
    Private dsCommand_CooperateOrganizationOpinion As SqlDataAdapter


    '定义查询命令
    Private GetCooperateOrganizationInfoCommand As SqlCommand
    Private GetCooperateOrganizationOpinionInfoCommand As SqlCommand


    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_CooperateOrganization = New SqlDataAdapter()
        dsCommand_CooperateOrganizationOpinion = New SqlDataAdapter()


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetCooperateOpinionInfo("null", "null")

    End Sub

    '获取合作单位意见信息
    Public Function GetCooperateOpinionInfo(ByVal strSQL_Condition_CooperateOrganization As String, ByVal strSQL_Condition_CooperateOrganizationOpinion As String) As DataSet

        Dim tempDs As New DataSet()

        If GetCooperateOrganizationInfoCommand Is Nothing Then

            GetCooperateOrganizationInfoCommand = New SqlCommand("GetCooperateOrganizationInfo", conn)
            GetCooperateOrganizationInfoCommand.CommandType = CommandType.StoredProcedure
            GetCooperateOrganizationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CooperateOrganization
            .SelectCommand = GetCooperateOrganizationInfoCommand
            .SelectCommand.Transaction = ts
            GetCooperateOrganizationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CooperateOrganization
            .Fill(tempDs, Table_Cooperate_Organization)
        End With

        If GetCooperateOrganizationOpinionInfoCommand Is Nothing Then

            GetCooperateOrganizationOpinionInfoCommand = New SqlCommand("GetCooperateOrganizationOpinionInfo", conn)
            GetCooperateOrganizationOpinionInfoCommand.CommandType = CommandType.StoredProcedure
            GetCooperateOrganizationOpinionInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CooperateOrganizationOpinion
            .SelectCommand = GetCooperateOrganizationOpinionInfoCommand
            .SelectCommand.Transaction = ts
            GetCooperateOrganizationOpinionInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CooperateOrganizationOpinion
            .Fill(tempDs, Table_Cooperate_Organization_Opinion)
        End With


        GetCooperateOpinionInfo = tempDs

    End Function

    '更新合作单位信息
    Private Function UpdateCooperateOrganization(ByVal CooperateOrganizationSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CooperateOrganization)

        With dsCommand_CooperateOrganization
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CooperateOrganizationSet, Table_Cooperate_Organization)

        End With

    End Function


    '更新合作单位意见信息
    Private Function UpdateCooperateOrganizationOpinion(ByVal CooperateOrganizationOpinionSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CooperateOrganizationOpinion)

        With dsCommand_CooperateOrganizationOpinion
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CooperateOrganizationOpinionSet, Table_Cooperate_Organization_Opinion)

        End With


    End Function

    '更新合作单位,意见信息
    Public Function UpdateCooperateOpinion(ByVal CooperateOpinionSet As DataSet)

        If CooperateOpinionSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If CooperateOpinionSet.HasChanges = False Then
            Exit Function
        End If

        UpdateCooperateOrganization(CooperateOpinionSet)
        UpdateCooperateOrganizationOpinion(CooperateOpinionSet)

        CooperateOpinionSet.AcceptChanges()

    End Function
End Class
