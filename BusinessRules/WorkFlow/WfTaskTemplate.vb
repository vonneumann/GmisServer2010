Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTaskTemplate
    Public Const Table_TaskTemplate As String = "task_template"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfTaskTemplate As SqlDataAdapter

    '定义查询命令
    Private GetWfTaskTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfTaskTemplate = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfTaskTemplateInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfTaskTemplateInfo(ByVal strSQL_Condition_WfTaskTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTaskTemplateInfoCommand Is Nothing Then

            GetWfTaskTemplateInfoCommand = New SqlCommand("GetWfTaskTemplateInfo", conn)
            GetWfTaskTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTaskTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTaskTemplate
            .SelectCommand = GetWfTaskTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTaskTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTaskTemplate
            .Fill(tempDs, Table_TaskTemplate)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTaskTemplate(ByVal WfTaskTemplateSet As DataSet)

        If WfTaskTemplateSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfTaskTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTaskTemplate)

        With dsCommand_WfTaskTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTaskTemplateSet, Table_TaskTemplate)

            WfTaskTemplateSet.AcceptChanges()
        End With

    End Function
End Class
