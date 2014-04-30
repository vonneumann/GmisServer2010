Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class JobPeriod
    Private Const Table_JobPeriod As String = "TJobPeriod"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_JobPeriod As SqlDataAdapter

    '定义查询命令
    Private GetJobPeriodCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_JobPeriod = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetJobPeriodInfo("null")
    End Sub

    '获取工作时间表信息
    Public Function GetJobPeriodInfo(ByVal strSQL_Condition_JobPeriod As String) As DataSet

        Dim tempDs As New DataSet()

        If GetJobPeriodCommand Is Nothing Then

            GetJobPeriodCommand = New SqlCommand("dbo.GetJobPeriodInfo", conn)
            GetJobPeriodCommand.CommandType = CommandType.StoredProcedure
            GetJobPeriodCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_JobPeriod
            .SelectCommand = GetJobPeriodCommand
            .SelectCommand.Transaction = ts
            GetJobPeriodCommand.Parameters("@Condition").Value = strSQL_Condition_JobPeriod
            .Fill(tempDs, Table_JobPeriod)
        End With

        Return tempDs

    End Function

    '更新假期信息
    Public Function UpdateJobPeriod(ByVal JobPeriodSet As DataSet)

        If JobPeriodSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If JobPeriodSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_JobPeriod)

        With dsCommand_JobPeriod
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(JobPeriodSet, Table_JobPeriod)

        End With

        JobPeriodSet.AcceptChanges()
    End Function
End Class
