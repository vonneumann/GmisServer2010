Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectPrompt
    Public Const Table_Project_Prompt As String = "project_prompt"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfProjectPrompt As SqlDataAdapter

    '定义查询命令
    Private GetWfTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WfProjectPrompt = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWfProjectPromptInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetWfProjectPromptInfo(ByVal strSQL_Condition_WfProjectPrompt As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTemplateInfoCommand Is Nothing Then

            GetWfTemplateInfoCommand = New SqlCommand("GetWfProjectPromptInfo", conn)
            GetWfTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectPrompt
            .SelectCommand = GetWfTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectPrompt
            .Fill(tempDs, Table_Project_Prompt)
        End With

        Return tempDs

    End Function

    '更新项目评价信息
    Public Function UpdateWfProjectPrompt(ByVal WfProjectPromptSet As DataSet)


        If WfProjectPromptSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WfProjectPromptSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectPrompt)

        With dsCommand_WfProjectPrompt
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectPromptSet, Table_Project_Prompt)

            WfProjectPromptSet.AcceptChanges()
        End With


    End Function
End Class
