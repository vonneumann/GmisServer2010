Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TracePlan
    Public Const Table_Trace_Plan As String = "trace_plan"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_TracePlan As SqlDataAdapter

    '定义查询命令
    Private GetTracePlanInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_TracePlan = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetTracePlanInfo("null")
    End Sub

    '获取检查记录基本信息
    Public Function GetTracePlanInfo(ByVal strSQL_Condition_TracePlan As String) As DataSet

        Dim tempDs As New DataSet()

        If GetTracePlanInfoCommand Is Nothing Then

            GetTracePlanInfoCommand = New SqlCommand("GetTracePlanInfo", conn)
            GetTracePlanInfoCommand.CommandType = CommandType.StoredProcedure
            GetTracePlanInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TracePlan
            .SelectCommand = GetTracePlanInfoCommand
            .SelectCommand.Transaction = ts
            GetTracePlanInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TracePlan
            .Fill(tempDs, Table_Trace_Plan)
        End With

        Return tempDs

    End Function

    '更新检查记录基本信息
    Public Function UpdateTracePlan(ByVal TracePlanSet As DataSet)

        If TracePlanSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If TracePlanSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TracePlan)

        With dsCommand_TracePlan
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TracePlanSet, Table_Trace_Plan)

        End With

        TracePlanSet.AcceptChanges()

    End Function
End Class
