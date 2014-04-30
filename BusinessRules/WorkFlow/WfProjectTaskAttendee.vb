Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTaskAttendee
    Public Const Table_Project_Task_Attendee As String = "project_task_attendee"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectTaskAttendee As SqlDataAdapter

    '定义查询命令
    Private GetWfProjectTaskAttendeeInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectTaskAttendee = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectTaskAttendeeInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfProjectTaskAttendeeInfo(ByVal strSQL_Condition_WfProjectTaskAttendee As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTaskAttendeeInfoCommand Is Nothing Then

            GetWfProjectTaskAttendeeInfoCommand = New SqlCommand("GetWfProjectTaskAttendeeInfo", conn)
            GetWfProjectTaskAttendeeInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTaskAttendeeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTaskAttendee
            .SelectCommand = GetWfProjectTaskAttendeeInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTaskAttendeeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTaskAttendee
            .Fill(tempDs, Table_Project_Task_Attendee)
        End With

        Return tempDs
      
    End Function

    '更新项目评价信息
    Public Function UpdateWfProjectTaskAttendee(ByVal WfProjectTaskAttendeeSet As DataSet)


        If WfProjectTaskAttendeeSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfProjectTaskAttendeeSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTaskAttendee)

        With dsCommand_WfProjectTaskAttendee
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTaskAttendeeSet, Table_Project_Task_Attendee)

            WfProjectTaskAttendeeSet.AcceptChanges()

        End With
    

    End Function

End Class
