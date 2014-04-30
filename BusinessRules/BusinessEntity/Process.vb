Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Process

    Public Const Table_Process As String = "process"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Process As SqlDataAdapter

    '定义查询命令
    Private GetProcessInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Process = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProcessInfo("null")
    End Sub

    '获取项目进度信息
    Public Function GetProcessInfo(ByVal strSQL_Condition_Process As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProcessInfoCommand Is Nothing Then

            GetProcessInfoCommand = New SqlCommand("GetProcessInfo", conn)
            GetProcessInfoCommand.CommandType = CommandType.StoredProcedure
            GetProcessInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Process
            .SelectCommand = GetProcessInfoCommand
            .SelectCommand.Transaction = ts
            GetProcessInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Process
            .Fill(tempDs, Table_Process)
        End With

        Return tempDs

    End Function

    '更新项目进度信息
    Public Function UpdateProcess(ByVal ProcessSet As DataSet)

        If ProcessSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If ProcessSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Process)

        With dsCommand_Process
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProcessSet, Table_Process)

        End With

        ProcessSet.AcceptChanges()
    End Function
End Class
