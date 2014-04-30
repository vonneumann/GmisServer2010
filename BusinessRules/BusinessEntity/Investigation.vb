Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Investigation

    Public Const Table_Investigation As String = "project_investigation"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Investigation As SqlDataAdapter

    '�����ѯ����
    Private GetInvestigationInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Investigation = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetInvestigationInfo("null")
    End Sub

    '��ȡ��ǰ�����¼��Ϣ
    Public Function GetInvestigationInfo(ByVal strSQL_Condition_Investigation As String) As DataSet

        Dim tempDs As New DataSet()

        If GetInvestigationInfoCommand Is Nothing Then

            GetInvestigationInfoCommand = New SqlCommand("GetInvestigationInfo", conn)
            GetInvestigationInfoCommand.CommandType = CommandType.StoredProcedure
            GetInvestigationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Investigation
            .SelectCommand = GetInvestigationInfoCommand
            .SelectCommand.Transaction = ts
            GetInvestigationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Investigation
            .Fill(tempDs, Table_Investigation)
        End With

        Return tempDs

    End Function

    '���±�ǰ�����¼��Ϣ
    Public Function UpdateInvestigation(ByVal InvestigationSet As DataSet)

        If InvestigationSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If InvestigationSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Investigation)

        With dsCommand_Investigation
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(InvestigationSet, Table_Investigation)

        End With

        InvestigationSet.AcceptChanges()

    End Function

End Class
