Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Project

    Public Const Table_Project As String = "project"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Project As SqlDataAdapter

    '定义查询命令
    Private GetProjectInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Project = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectInfo("null")
    End Sub

    '获取项目基本信息
    Public Function GetProjectInfo(ByVal strSQL_Condition_Project As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectInfoCommand Is Nothing Then

            GetProjectInfoCommand = New SqlCommand("GetProjectInfo", conn)
            GetProjectInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Project
            .SelectCommand = GetProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Project
            .Fill(tempDs, Table_Project)
        End With

        Return tempDs

    End Function

    '更新项目基本信息
    Public Function UpdateProject(ByVal ProjectSet As DataSet)

        If ProjectSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Project)

        With dsCommand_Project
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectSet, Table_Project)

        End With

        ProjectSet.AcceptChanges()

    End Function
End Class
