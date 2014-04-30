Option Explicit On 


Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CommitteemanOpinion
    Public Const Table_Committeeman_Opinion As String = "committeeman_opinion"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_CommitteemanOpinion As SqlDataAdapter

    '定义查询命令
    Private GetCommitteemanOpinionInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction
    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_CommitteemanOpinion = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetCommitteemanOpinionInfo("null")
    End Sub

    '获取评委意见信息
    Public Function GetCommitteemanOpinionInfo(ByVal strSQL_Condition_CommitteemanOpinion As String) As DataSet

        Dim tempDs As New DataSet()

        If GetCommitteemanOpinionInfoCommand Is Nothing Then

            GetCommitteemanOpinionInfoCommand = New SqlCommand("GetCommitteemanOpinionInfo", conn)
            GetCommitteemanOpinionInfoCommand.CommandType = CommandType.StoredProcedure
            GetCommitteemanOpinionInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CommitteemanOpinion
            .SelectCommand = GetCommitteemanOpinionInfoCommand
            .SelectCommand.Transaction = ts
            GetCommitteemanOpinionInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CommitteemanOpinion
            .Fill(tempDs, Table_Committeeman_Opinion)
        End With

        Return tempDs

    End Function

    '更新评委意见信息
    Public Function UpdateCommitteemanOpinion(ByVal CommitteemanOpinionSet As DataSet)

        If CommitteemanOpinionSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If CommitteemanOpinionSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CommitteemanOpinion)

        With dsCommand_CommitteemanOpinion
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CommitteemanOpinionSet, Table_Committeeman_Opinion)

            CommitteemanOpinionSet.AcceptChanges()
        End With


    End Function
End Class
