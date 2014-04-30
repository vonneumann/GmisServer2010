Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ConfTrial

    Public Const Table_Conference_trial As String = "conference_trial"
    Public Const Table_Committeeman_opinion As String = "committeeman_opinion"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ConferenceTrial As SqlDataAdapter
    Private dsCommand_CommitteemanOpinion As SqlDataAdapter

    '�����ѯ����
    Private GetConferenceTrialInfoCommand As SqlCommand
    Private GetCommitteemanOpinionInfoCommand As SqlCommand


    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ConferenceTrial = New SqlDataAdapter()
        dsCommand_CommitteemanOpinion = New SqlDataAdapter()


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetConfTrialInfo("null", "null")
    End Sub

    '��ȡ���������Ϣ
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

    '�������������Ϣ
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

    '������ί�����Ϣ
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

    '�����������,��ί�����Ϣ
    Public Function UpdateConfTrial(ByVal ConfTrialSet As DataSet)

        If ConfTrialSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If ConfTrialSet.HasChanges = False Then
            Exit Function
        End If


        'ɾ������
        If IsNothing(ConfTrialSet.GetChanges(DataRowState.Deleted)) = False Then
            '��ɾ��ϸ����ɾ����
            UpdatedsCommitteemanOpinion(ConfTrialSet.GetChanges(DataRowState.Deleted))
            UpdateConferenceTrial(ConfTrialSet.GetChanges(DataRowState.Deleted))

        End If

        '��������
        If IsNothing(ConfTrialSet.GetChanges(DataRowState.Added)) = False Then

            UpdateConferenceTrial(ConfTrialSet.GetChanges(DataRowState.Added))
            UpdatedsCommitteemanOpinion(ConfTrialSet.GetChanges(DataRowState.Added))
        End If

        '���²���
        If IsNothing(ConfTrialSet.GetChanges(DataRowState.Modified)) = False Then

            UpdateConferenceTrial(ConfTrialSet.GetChanges(DataRowState.Modified))
            UpdatedsCommitteemanOpinion(ConfTrialSet.GetChanges(DataRowState.Modified))
        End If

        ConfTrialSet.AcceptChanges()
    End Function
End Class
