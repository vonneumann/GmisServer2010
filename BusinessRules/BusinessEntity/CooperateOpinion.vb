Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CooperateOpinion

    Public Const Table_Cooperate_Organization As String = "cooperate_organization"
    Public Const Table_Cooperate_Organization_Opinion As String = "cooperate_organization_opinion"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_CooperateOrganization As SqlDataAdapter
    Private dsCommand_CooperateOrganizationOpinion As SqlDataAdapter


    '�����ѯ����
    Private GetCooperateOrganizationInfoCommand As SqlCommand
    Private GetCooperateOrganizationOpinionInfoCommand As SqlCommand


    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_CooperateOrganization = New SqlDataAdapter()
        dsCommand_CooperateOrganizationOpinion = New SqlDataAdapter()


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetCooperateOpinionInfo("null", "null")

    End Sub

    '��ȡ������λ�����Ϣ
    Public Function GetCooperateOpinionInfo(ByVal strSQL_Condition_CooperateOrganization As String, ByVal strSQL_Condition_CooperateOrganizationOpinion As String) As DataSet

        Dim tempDs As New DataSet()

        If GetCooperateOrganizationInfoCommand Is Nothing Then

            GetCooperateOrganizationInfoCommand = New SqlCommand("GetCooperateOrganizationInfo", conn)
            GetCooperateOrganizationInfoCommand.CommandType = CommandType.StoredProcedure
            GetCooperateOrganizationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CooperateOrganization
            .SelectCommand = GetCooperateOrganizationInfoCommand
            .SelectCommand.Transaction = ts
            GetCooperateOrganizationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CooperateOrganization
            .Fill(tempDs, Table_Cooperate_Organization)
        End With

        If GetCooperateOrganizationOpinionInfoCommand Is Nothing Then

            GetCooperateOrganizationOpinionInfoCommand = New SqlCommand("GetCooperateOrganizationOpinionInfo", conn)
            GetCooperateOrganizationOpinionInfoCommand.CommandType = CommandType.StoredProcedure
            GetCooperateOrganizationOpinionInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CooperateOrganizationOpinion
            .SelectCommand = GetCooperateOrganizationOpinionInfoCommand
            .SelectCommand.Transaction = ts
            GetCooperateOrganizationOpinionInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CooperateOrganizationOpinion
            .Fill(tempDs, Table_Cooperate_Organization_Opinion)
        End With


        GetCooperateOpinionInfo = tempDs

    End Function

    '���º�����λ��Ϣ
    Private Function UpdateCooperateOrganization(ByVal CooperateOrganizationSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CooperateOrganization)

        With dsCommand_CooperateOrganization
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CooperateOrganizationSet, Table_Cooperate_Organization)

        End With

    End Function


    '���º�����λ�����Ϣ
    Private Function UpdateCooperateOrganizationOpinion(ByVal CooperateOrganizationOpinionSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CooperateOrganizationOpinion)

        With dsCommand_CooperateOrganizationOpinion
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CooperateOrganizationOpinionSet, Table_Cooperate_Organization_Opinion)

        End With


    End Function

    '���º�����λ,�����Ϣ
    Public Function UpdateCooperateOpinion(ByVal CooperateOpinionSet As DataSet)

        If CooperateOpinionSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If CooperateOpinionSet.HasChanges = False Then
            Exit Function
        End If

        UpdateCooperateOrganization(CooperateOpinionSet)
        UpdateCooperateOrganizationOpinion(CooperateOpinionSet)

        CooperateOpinionSet.AcceptChanges()

    End Function
End Class
