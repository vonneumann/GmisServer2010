Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WorkflowType
    Public Const Table_WorkflowType As String = "workflow"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WorkflowType As SqlDataAdapter

    '定义查询命令
    Private GetWorkflowTypeInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WorkflowType = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWorkflowTypeInfo("null")
    End Sub

    '获取工作流类型信息
    Public Function GetWorkflowTypeInfo(ByVal strSQL_Condition_WorkflowType As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWorkflowTypeInfoCommand Is Nothing Then

            GetWorkflowTypeInfoCommand = New SqlCommand("GetWorkflowTypeInfo", conn)
            GetWorkflowTypeInfoCommand.CommandType = CommandType.StoredProcedure
            GetWorkflowTypeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WorkflowType
            .SelectCommand = GetWorkflowTypeInfoCommand
            .SelectCommand.Transaction = ts
            GetWorkflowTypeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WorkflowType
            .Fill(tempDs, Table_WorkflowType)
        End With

        Return tempDs


    End Function

    '更新工作流类型信息
    Public Function UpdateWorkflowType(ByVal WorkflowTypeSet As DataSet)

        '如果记录集未发生任何变化，则退出过程
        If WorkflowTypeSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WorkflowType)

        With dsCommand_WorkflowType
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WorkflowTypeSet, Table_WorkflowType)

        End With

        WorkflowTypeSet.AcceptChanges()
    End Function
End Class
