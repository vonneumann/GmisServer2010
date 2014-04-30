Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectEndCase

    Public Const Table_ProjectEndCase As String = "project_end_case"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectEndCase As SqlDataAdapter

    '定义查询命令
    Private GetProjectEndCaseInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectEndCase = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectEndCaseInfo("null")

    End Sub

    '获取工作记录信息
    Public Function GetProjectEndCaseInfo(ByVal strSQL_Condition_WorkLog As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectEndCaseInfoCommand Is Nothing Then

            GetProjectEndCaseInfoCommand = New SqlCommand("GetProjectEndCaseInfo", conn)
            GetProjectEndCaseInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectEndCaseInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectEndCase
            .SelectCommand = GetProjectEndCaseInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectEndCaseInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WorkLog
            .Fill(tempDs, Table_ProjectEndCase)
        End With

        Return tempDs


    End Function

    '更新工作记录信息
    Public Function UpdateProjectEndCase(ByVal ProjectEndCase As DataSet)

        If ProjectEndCase Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectEndCase.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectEndCase)

        With dsCommand_ProjectEndCase
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectEndCase, Table_ProjectEndCase)

            ProjectEndCase.AcceptChanges()
        End With


    End Function

End Class
