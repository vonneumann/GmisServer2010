
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTaskTransferTemplate
    Public Const Table_Task_Transfer_Template As String = "task_transfer_template"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfTaskTransferTemplate As SqlDataAdapter

    '定义查询命令
    Private GetWfTaskTransferTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfTaskTransferTemplate = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfTaskTransferTemplateInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfTaskTransferTemplateInfo(ByVal strSQL_Condition_WfTaskTransferTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTaskTransferTemplateInfoCommand Is Nothing Then

            GetWfTaskTransferTemplateInfoCommand = New SqlCommand("GetWfTaskTransferTemplateInfo", conn)
            GetWfTaskTransferTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTaskTransferTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTaskTransferTemplate
            .SelectCommand = GetWfTaskTransferTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTaskTransferTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTaskTransferTemplate
            .Fill(tempDs, Table_Task_Transfer_Template)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTaskTransferTemplate(ByVal WfTaskTransferTemplateSet As DataSet)

        If WfTaskTransferTemplateSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfTaskTransferTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTaskTransferTemplate)

        With dsCommand_WfTaskTransferTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTaskTransferTemplateSet, Table_Task_Transfer_Template)

            WfTaskTransferTemplateSet.AcceptChanges()
        End With

    End Function
End Class
