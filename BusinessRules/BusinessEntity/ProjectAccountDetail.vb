Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectAccountDetail

    Public Const Table_Project_Account_Detail As String = "project_account_detail"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ProjectAccountDetail As SqlDataAdapter

    '定义查询命令
    Private GetProjectAccountDetailInfoCommand As SqlCommand
    Private GetMaxProjectAccountDetailNumCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ProjectAccountDetail = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetProjectAccountDetailInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetProjectAccountDetailInfo(ByVal strSQL_Condition_ProjectAccountDetail As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectAccountDetailInfoCommand Is Nothing Then

            GetProjectAccountDetailInfoCommand = New SqlCommand("GetProjectAccountDetailInfo", conn)
            GetProjectAccountDetailInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectAccountDetailInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectAccountDetail
            .SelectCommand = GetProjectAccountDetailInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectAccountDetailInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectAccountDetail
            .Fill(tempDs, Table_Project_Account_Detail)
        End With

        Return tempDs

    End Function


    '获取最大序列号
    Public Function GetMaxProjectAccountDetailNum(ByVal projectID As String) As Integer

        If GetMaxProjectAccountDetailNumCommand Is Nothing Then

            GetMaxProjectAccountDetailNumCommand = New SqlCommand("GetMaxProjectAccountDetailNum", conn)
            GetMaxProjectAccountDetailNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxProjectAccountDetailNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxProjectAccountDetailNumCommand.Parameters.Add(New SqlParameter("@maxProjectAccountDetailNum", SqlDbType.Int))
            GetMaxProjectAccountDetailNumCommand.Parameters.Item("@maxProjectAccountDetailNum").Direction = ParameterDirection.Output
            GetMaxProjectAccountDetailNumCommand.Transaction = ts
        End If

        GetMaxProjectAccountDetailNumCommand.Parameters("@projectID").Value = projectID
        GetMaxProjectAccountDetailNumCommand.ExecuteNonQuery()
        GetMaxProjectAccountDetailNum = GetMaxProjectAccountDetailNumCommand.Parameters.Item("@maxProjectAccountDetailNum").Value
    End Function

    '更新项目评价信息
    Public Function UpdateProjectAccountDetail(ByVal ProjectAccountDetailSet As DataSet)

        If ProjectAccountDetailSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ProjectAccountDetailSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectAccountDetail)

        With dsCommand_ProjectAccountDetail
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectAccountDetailSet, Table_Project_Account_Detail)

            ProjectAccountDetailSet.AcceptChanges()
        End With


    End Function
End Class
