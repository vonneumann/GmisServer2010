
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTaskRoleTemplate
    Public Const Table_Task_Role_Template As String = "task_role_template"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfTaskRoleTemplate As SqlDataAdapter

    '定义查询命令
    Private GetWfTaskRoleTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfTaskRoleTemplate = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfTaskRoleTemplateInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfTaskRoleTemplateInfo(ByVal strSQL_Condition_WfTaskRoleTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTaskRoleTemplateInfoCommand Is Nothing Then

            GetWfTaskRoleTemplateInfoCommand = New SqlCommand("GetWfTaskRoleTemplateInfo", conn)
            GetWfTaskRoleTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTaskRoleTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTaskRoleTemplate
            .SelectCommand = GetWfTaskRoleTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTaskRoleTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTaskRoleTemplate
            .Fill(tempDs, Table_Task_Role_Template)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTaskRoleTemplate(ByVal WfTaskRoleTemplateSet As DataSet)

        If WfTaskRoleTemplateSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfTaskRoleTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTaskRoleTemplate)

        With dsCommand_WfTaskRoleTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTaskRoleTemplateSet, Table_Task_Role_Template)

            WfTaskRoleTemplateSet.AcceptChanges()
        End With


    End Function
End Class
