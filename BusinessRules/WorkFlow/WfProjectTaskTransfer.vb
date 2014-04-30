Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTaskTransfer
    Public Const Table_Project_Task_Transfer As String = "project_task_transfer"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectTaskTransfer As SqlDataAdapter

    '定义查询命令
    Private GetWfProjectTaskTransferInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectTaskTransfer = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectTaskTransferInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfProjectTaskTransferInfo(ByVal strSQL_Condition_WfProjectTaskTransfer As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTaskTransferInfoCommand Is Nothing Then

            GetWfProjectTaskTransferInfoCommand = New SqlCommand("GetWfProjectTaskTransferInfo", conn)
            GetWfProjectTaskTransferInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTaskTransferInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTaskTransfer
            .SelectCommand = GetWfProjectTaskTransferInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTaskTransferInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTaskTransfer
            .Fill(tempDs, Table_Project_Task_Transfer)
        End With

        Return tempDs

    End Function

    '更新项目评价信息
    Public Function UpdateWfProjectTaskTransfer(ByVal WfProjectTaskTransferSet As DataSet)

        If WfProjectTaskTransferSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If WfProjectTaskTransferSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTaskTransfer)

        With dsCommand_WfProjectTaskTransfer
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTaskTransferSet, Table_Project_Task_Transfer)

            WfProjectTaskTransferSet.AcceptChanges()
        End With

    End Function
End Class
