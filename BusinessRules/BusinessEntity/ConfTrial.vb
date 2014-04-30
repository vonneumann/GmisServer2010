Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ConfTrial

    Public Const Table_Conference_trial As String = "conference_trial"
    Public Const Table_Committeeman_opinion As String = "committeeman_opinion"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ConferenceTrial As SqlDataAdapter
    Private dsCommand_CommitteemanOpinion As SqlDataAdapter

    '定义查询命令
    Private GetConferenceTrialInfoCommand As SqlCommand
    Private GetCommitteemanOpinionInfoCommand As SqlCommand


    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ConferenceTrial = New SqlDataAdapter()
        dsCommand_CommitteemanOpinion = New SqlDataAdapter()


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetConfTrialInfo("null", "null")
    End Sub

    '获取评审意见信息
    Public Function GetConfTrialInfo(ByVal strSQL_Condition_ConferenceTrial As String, ByVal strSQL_Condition_CommitteemanOpinion As String) As DataSet

        Dim tempDs As New DataSet()

        If GetConferenceTrialInfoCommand Is Nothing Then

            GetConferenceTrialInfoCommand = New SqlCommand("GetConferenceTrialInfo", conn)
            GetConferenceTrialInfoCommand.CommandType = CommandType.StoredProcedure
            GetConferenceTrialInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ConferenceTrial
            .SelectCommand = GetConferenceTrialInfoCommand
            .SelectCommand.Transaction = ts
            GetConferenceTrialInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ConferenceTrial
            .Fill(tempDs, Table_Conference_trial)
        End With

        If GetCommitteemanOpinionInfoCommand Is Nothing Then

            GetCommitteemanOpinionInfoCommand = New SqlCommand("GetCommitteemanOpinionInfo", conn)
            GetCommitteemanOpinionInfoCommand.CommandType = CommandType.StoredProcedure
            GetCommitteemanOpinionInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CommitteemanOpinion
            .SelectCommand = GetCommitteemanOpinionInfoCommand
            .SelectCommand.Transaction = ts
            GetCommitteemanOpinionInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CommitteemanOpinion
            .Fill(tempDs, Table_Committeeman_opinion)
        End With

        GetConfTrialInfo = tempDs

    End Function

    '更新评审意见信息
    Private Function UpdateConferenceTrial(ByVal ConferenceTrialSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ConferenceTrial)

        With dsCommand_ConferenceTrial
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ConferenceTrialSet, Table_Conference_trial)

        End With

    End Function

    '更新评委意见信息
    Private Function UpdatedsCommitteemanOpinion(ByVal CommitteemanOpinionSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CommitteemanOpinion)

        With dsCommand_CommitteemanOpinion
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CommitteemanOpinionSet, Table_Committeeman_opinion)

        End With

    End Function

    '更新评审意见,评委意见信息
    Public Function UpdateConfTrial(ByVal ConfTrialSet As DataSet)

        If ConfTrialSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If ConfTrialSet.HasChanges = False Then
            Exit Function
        End If


        '删除操作
        If IsNothing(ConfTrialSet.GetChanges(DataRowState.Deleted)) = False Then
            '先删明细表，再删主表
            UpdatedsCommitteemanOpinion(ConfTrialSet.GetChanges(DataRowState.Deleted))
            UpdateConferenceTrial(ConfTrialSet.GetChanges(DataRowState.Deleted))

        End If

        '新增操作
        If IsNothing(ConfTrialSet.GetChanges(DataRowState.Added)) = False Then

            UpdateConferenceTrial(ConfTrialSet.GetChanges(DataRowState.Added))
            UpdatedsCommitteemanOpinion(ConfTrialSet.GetChanges(DataRowState.Added))
        End If

        '更新操作
        If IsNothing(ConfTrialSet.GetChanges(DataRowState.Modified)) = False Then

            UpdateConferenceTrial(ConfTrialSet.GetChanges(DataRowState.Modified))
            UpdatedsCommitteemanOpinion(ConfTrialSet.GetChanges(DataRowState.Modified))
        End If

        ConfTrialSet.AcceptChanges()
    End Function
End Class
