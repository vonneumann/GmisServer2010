Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TerminateReport

    Public Const Table_TerminateReport As String = "project_terminate_report"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_TerminateReport As SqlDataAdapter

    '定义查询命令
    Private GetTerminateReportInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_TerminateReport = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetTerminateReportInfo("null")
    End Sub

    '获取项目终止报告信息
    Public Function GetTerminateReportInfo(ByVal strSQL_Condition_TerminateReport As String) As DataSet

        Dim tempDs As New DataSet()

        If GetTerminateReportInfoCommand Is Nothing Then

            GetTerminateReportInfoCommand = New SqlCommand("GetTerminateReportInfo", conn)
            GetTerminateReportInfoCommand.CommandType = CommandType.StoredProcedure
            GetTerminateReportInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TerminateReport
            .SelectCommand = GetTerminateReportInfoCommand
            .SelectCommand.Transaction = ts
            GetTerminateReportInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TerminateReport
            .Fill(tempDs, Table_TerminateReport)
        End With

        Return tempDs

    End Function

    '更新项目终止报告信息
    Public Function UpdateTerminateReport(ByVal TerminateReportSet As DataSet)

        If TerminateReportSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If TerminateReportSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TerminateReport)

        With dsCommand_TerminateReport
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TerminateReportSet, Table_TerminateReport)

        End With

        TerminateReportSet.AcceptChanges()

    End Function
End Class
