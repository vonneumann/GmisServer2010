Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTimingTaskTemplate
    Public Const Table_Timing_Task_Template As String = "timing_task_template"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfTimingTaskTemplate As SqlDataAdapter

    '定义查询命令
    Private GetWfTimingTaskTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfTimingTaskTemplate = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfTimingTaskTemplateInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfTimingTaskTemplateInfo(ByVal strSQL_Condition_WfTimingTaskTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTimingTaskTemplateInfoCommand Is Nothing Then

            GetWfTimingTaskTemplateInfoCommand = New SqlCommand("GetWfTimingTaskTemplateInfo", conn)
            GetWfTimingTaskTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTimingTaskTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTimingTaskTemplate
            .SelectCommand = GetWfTimingTaskTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTimingTaskTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTimingTaskTemplate
            .Fill(tempDs, Table_Timing_Task_Template)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTimingTaskTemplate(ByVal WfTimingTaskTemplateSet As DataSet)

        If WfTimingTaskTemplateSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfTimingTaskTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTimingTaskTemplate)

        With dsCommand_WfTimingTaskTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTimingTaskTemplateSet, Table_Timing_Task_Template)

            WfTimingTaskTemplateSet.AcceptChanges()
        End With


    End Function

End Class
