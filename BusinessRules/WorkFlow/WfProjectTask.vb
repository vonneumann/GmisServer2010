Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTask

    Public Const Table_Project_Task As String = "project_task"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectTask As SqlDataAdapter

    '定义查询命令
    Private GetWfTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectTask = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectTaskInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfProjectTaskInfo(ByVal strSQL_Condition_WfProjectTask As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTemplateInfoCommand Is Nothing Then

            GetWfTemplateInfoCommand = New SqlCommand("GetWfProjectTaskInfo", conn)
            GetWfTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTask
            .SelectCommand = GetWfTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTask
            .Fill(tempDs, Table_Project_Task)
        End With
        Return tempDs
      
    End Function

    '更新项目评价信息
    Public Function UpdateWfProjectTask(ByVal WfProjectTaskSet As DataSet)

        If WfProjectTaskSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfProjectTaskSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTask)

        With dsCommand_WfProjectTask
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTaskSet, Table_Project_Task)

            WfProjectTaskSet.AcceptChanges()
        End With


    End Function



End Class
