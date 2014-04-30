Option Explicit On 


Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CommitteemanOpinion
    Public Const Table_Committeeman_Opinion As String = "committeeman_opinion"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_CommitteemanOpinion As SqlDataAdapter

    '�����ѯ����
    Private GetCommitteemanOpinionInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction
    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_CommitteemanOpinion = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetCommitteemanOpinionInfo("null")
    End Sub

    '��ȡ��ί�����Ϣ
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

    '������ί�����Ϣ
    Public Function UpdateCommitteemanOpinion(ByVal CommitteemanOpinionSet As DataSet)

        If CommitteemanOpinionSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
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
