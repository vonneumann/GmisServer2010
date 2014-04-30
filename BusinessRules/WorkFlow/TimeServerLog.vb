Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TimeServerLog
    Public Const Table_Project_Task_Attendee As String = "time_server_log"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_TimeServerLog As SqlDataAdapter

    '定义查询命令
    Private GetTimeServerLogInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_TimeServerLog = New SqlDataAdapter

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetTimeServerLogInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetTimeServerLogInfo(ByVal strSQL_Condition_TimeServerLog As String) As DataSet

        Dim tempDs As New DataSet

        If GetTimeServerLogInfoCommand Is Nothing Then

            GetTimeServerLogInfoCommand = New SqlCommand("GetTimeServerLogInfo", conn)
            GetTimeServerLogInfoCommand.CommandType = CommandType.StoredProcedure
            GetTimeServerLogInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TimeServerLog
            .SelectCommand = GetTimeServerLogInfoCommand
            .SelectCommand.Transaction = ts
            GetTimeServerLogInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TimeServerLog
            .Fill(tempDs, Table_Project_Task_Attendee)
        End With

        Return tempDs

    End Function

    '更新项目评价信息
    Public Function UpdateTimeServerLog(ByVal TimeServerLogSet As DataSet)


        If TimeServerLogSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If TimeServerLogSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TimeServerLog)

        With dsCommand_TimeServerLog
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TimeServerLogSet, Table_Project_Task_Attendee)

            TimeServerLogSet.AcceptChanges()

        End With


    End Function

End Class
