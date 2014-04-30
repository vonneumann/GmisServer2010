Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'岗位和工作职责
Public Class PostAndResponsibility
    Private Const Table_Post As String = "TPost"
    Private Const Table_JobResponsibility As String = "TJobResponsibility"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Post As SqlDataAdapter
    Private dsCommand_JobResponsibility As SqlDataAdapter

    '定义查询命令
    Private GetPostCommand As SqlCommand
    Private GetJobResponsibilityCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Post = New SqlDataAdapter()
        dsCommand_JobResponsibility = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetPostAndJobResponsibilityInfo("null", "null")
    End Sub

    '获取岗位信息
    Public Function GetPostAndJobResponsibilityInfo(ByVal strSQL_Condition_Post As String, ByVal strSQL_Condition_JobResponsibility As String) As DataSet

        Dim tempDs As New DataSet()

        If GetPostCommand Is Nothing Then

            GetPostCommand = New SqlCommand("dbo.GetPostInfo", conn)
            GetPostCommand.CommandType = CommandType.StoredProcedure
            GetPostCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Post
            .SelectCommand = GetPostCommand
            .SelectCommand.Transaction = ts
            GetPostCommand.Parameters("@Condition").Value = strSQL_Condition_Post
            .Fill(tempDs, Table_Post)
        End With

        If GetJobResponsibilityCommand Is Nothing Then
            GetJobResponsibilityCommand = New SqlCommand("dbo.GetJobResponsibility", conn)
            GetJobResponsibilityCommand.CommandType = CommandType.StoredProcedure
            GetJobResponsibilityCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))
        End If

        With dsCommand_JobResponsibility
            .SelectCommand = GetJobResponsibilityCommand
            .SelectCommand.Transaction = ts
            .SelectCommand.Parameters("@Condition").Value = strSQL_Condition_JobResponsibility
            .Fill(tempDs, Table_JobResponsibility)
        End With

        Return tempDs

    End Function

    Private Function UpdatePost(ByVal PostDataset As DataSet)
        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Post)

        With dsCommand_Post

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(PostDataset, Table_Post)

        End With
    End Function


    Private Function UpdateJobResponsibility(ByVal JobResponsibilitySet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_JobResponsibility)

        With dsCommand_JobResponsibility

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(JobResponsibilitySet, Table_JobResponsibility)

        End With
    End Function

    Public Function UpdatePostAndJobResponsibility(ByVal commitSet As DataSet)
        If commitSet Is Nothing Then
            Exit Function
        End If

        If Not commitSet.HasChanges Then
            Exit Function
        End If


        '删除操作
        If IsNothing(commitSet.GetChanges(DataRowState.Deleted)) = False Then
            '先删明细表，再删主表
            UpdatePost(commitSet.GetChanges(DataRowState.Deleted))
            UpdateJobResponsibility(commitSet.GetChanges(DataRowState.Deleted))

        End If

        '新增作
        If IsNothing(commitSet.GetChanges(DataRowState.Added)) = False Then
            UpdatePost(commitSet.GetChanges(DataRowState.Added))
            UpdateJobResponsibility(commitSet.GetChanges(DataRowState.Added))
        End If

        '更新操作
        If IsNothing(commitSet.GetChanges(DataRowState.Modified)) = False Then
            UpdatePost(commitSet.GetChanges(DataRowState.Modified))
            UpdateJobResponsibility(commitSet.GetChanges(DataRowState.Modified))
        End If

        commitSet.AcceptChanges()

    End Function



End Class
