Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class UserPost
    Private Const Table_UserPost As String = "TUserPost"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_UserPost As SqlDataAdapter

    '定义查询命令
    Private GetUserPostCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_UserPost = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetUserPostInfo("null")
    End Sub

    '获取假期信息
    Public Function GetUserPostInfo(ByVal strSQL_Condition_UserPost As String) As DataSet

        Dim tempDs As New DataSet()

        If GetUserPostCommand Is Nothing Then

            GetUserPostCommand = New SqlCommand("dbo.GetUserPostInfo", conn)
            GetUserPostCommand.CommandType = CommandType.StoredProcedure
            GetUserPostCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_UserPost
            .SelectCommand = GetUserPostCommand
            .SelectCommand.Transaction = ts
            GetUserPostCommand.Parameters("@Condition").Value = strSQL_Condition_UserPost
            .Fill(tempDs, Table_UserPost)
        End With

        Return tempDs

    End Function

    '更新假期信息
    Public Function UpdateUserPost(ByVal UserPostSet As DataSet)

        If UserPostSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If UserPostSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_UserPost)

        With dsCommand_UserPost
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(UserPostSet, Table_UserPost)

        End With

        UserPostSet.AcceptChanges()
    End Function
End Class
