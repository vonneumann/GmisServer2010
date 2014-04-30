Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectMessages

    Public Const Table_Project_Messages As String = "project_task"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectMessages As SqlDataAdapter

    '定义查询命令
    Private GetWfProjectMessagesInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectMessages = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectMessagesInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfProjectMessagesInfo(ByVal strSQL_Condition_WfProjectMessages As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectMessagesInfoCommand Is Nothing Then

            GetWfProjectMessagesInfoCommand = New SqlCommand("GetWfProjectMessagesInfo", conn)
            GetWfProjectMessagesInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectMessagesInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectMessages
            .SelectCommand = GetWfProjectMessagesInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectMessagesInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectMessages
            .Fill(tempDs, Table_Project_Messages)
        End With

        Return tempDs

    End Function

    '更新项目评价信息
    Public Function UpdateWfProjectMessages(ByVal WfProjectMessagesSet As DataSet)

        If WfProjectMessagesSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfProjectMessagesSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectMessages)

        With dsCommand_WfProjectMessages
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectMessagesSet, Table_Project_Messages)

            WfProjectMessagesSet.AcceptChanges()
        End With
    End Function
End Class
