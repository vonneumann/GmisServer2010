Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTimingTask

    Public Const Table_Project_Timing_Task As String = "project_timing_task"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectTimingTask As SqlDataAdapter

    '定义查询命令
    Private GetWfProjectTimingTaskInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectTimingTask = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectTimingTaskInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfProjectTimingTaskInfo(ByVal strSQL_Condition_WfProjectTimingTask As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTimingTaskInfoCommand Is Nothing Then

            GetWfProjectTimingTaskInfoCommand = New SqlCommand("GetWfProjectTimingTaskInfo", conn)
            GetWfProjectTimingTaskInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTimingTaskInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTimingTask
            .SelectCommand = GetWfProjectTimingTaskInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTimingTaskInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTimingTask
            .Fill(tempDs, Table_Project_Timing_Task)
        End With

        Return tempDs

    End Function

    '更新项目评价信息
    Public Function UpdateWfProjectTimingTask(ByVal WfProjectTimingTaskSet As DataSet)

        If WfProjectTimingTaskSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If WfProjectTimingTaskSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTimingTask)

        With dsCommand_WfProjectTimingTask
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTimingTaskSet, Table_Project_Timing_Task)

            WfProjectTimingTaskSet.AcceptChanges()
        End With


    End Function
End Class
