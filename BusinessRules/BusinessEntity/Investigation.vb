Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Investigation

    Public Const Table_Investigation As String = "project_investigation"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Investigation As SqlDataAdapter

    '定义查询命令
    Private GetInvestigationInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Investigation = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetInvestigationInfo("null")
    End Sub

    '获取保前调查记录信息
    Public Function GetInvestigationInfo(ByVal strSQL_Condition_Investigation As String) As DataSet

        Dim tempDs As New DataSet()

        If GetInvestigationInfoCommand Is Nothing Then

            GetInvestigationInfoCommand = New SqlCommand("GetInvestigationInfo", conn)
            GetInvestigationInfoCommand.CommandType = CommandType.StoredProcedure
            GetInvestigationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Investigation
            .SelectCommand = GetInvestigationInfoCommand
            .SelectCommand.Transaction = ts
            GetInvestigationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Investigation
            .Fill(tempDs, Table_Investigation)
        End With

        Return tempDs

    End Function

    '更新保前调查记录信息
    Public Function UpdateInvestigation(ByVal InvestigationSet As DataSet)

        If InvestigationSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If InvestigationSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Investigation)

        With dsCommand_Investigation
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(InvestigationSet, Table_Investigation)

        End With

        InvestigationSet.AcceptChanges()

    End Function

End Class
