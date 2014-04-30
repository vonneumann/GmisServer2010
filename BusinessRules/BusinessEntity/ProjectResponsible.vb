Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectResponsible

    Public Const Table_ProjectResponsible As String = "project_responsible"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectResponsible As SqlDataAdapter

    '定义查询命令
    Private GetProjectResponsibleInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectResponsible = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectResponsibleInfo("null")

    End Sub

    '获取工作记录信息
    Public Function GetProjectResponsibleInfo(ByVal strSQL_Condition As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectResponsibleInfoCommand Is Nothing Then

            GetProjectResponsibleInfoCommand = New SqlCommand("GetProjectResponsible", conn)
            GetProjectResponsibleInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectResponsibleInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectResponsible
            .SelectCommand = GetProjectResponsibleInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectResponsibleInfoCommand.Parameters("@Condition").Value = strSQL_Condition
            .Fill(tempDs, Table_ProjectResponsible)
        End With

        Return tempDs


    End Function

    '更新工作记录信息
    Public Function UpdateProjectResponsible(ByVal ProjectResponsible As DataSet)

        If ProjectResponsible Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectResponsible.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectResponsible)

        With dsCommand_ProjectResponsible
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectResponsible, Table_ProjectResponsible)

            ProjectResponsible.AcceptChanges()
        End With


    End Function

End Class
