Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TerminateReport

    Public Const Table_TerminateReport As String = "project_terminate_report"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_TerminateReport As SqlDataAdapter

    '�����ѯ����
    Private GetTerminateReportInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_TerminateReport = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetTerminateReportInfo("null")
    End Sub

    '��ȡ��Ŀ��ֹ������Ϣ
    Public Function GetTerminateReportInfo(ByVal strSQL_Condition_TerminateReport As String) As DataSet

        Dim tempDs As New DataSet()

        If GetTerminateReportInfoCommand Is Nothing Then

            GetTerminateReportInfoCommand = New SqlCommand("GetTerminateReportInfo", conn)
            GetTerminateReportInfoCommand.CommandType = CommandType.StoredProcedure
            GetTerminateReportInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TerminateReport
            .SelectCommand = GetTerminateReportInfoCommand
            .SelectCommand.Transaction = ts
            GetTerminateReportInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TerminateReport
            .Fill(tempDs, Table_TerminateReport)
        End With

        Return tempDs

    End Function

    '������Ŀ��ֹ������Ϣ
    Public Function UpdateTerminateReport(ByVal TerminateReportSet As DataSet)

        If TerminateReportSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If TerminateReportSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TerminateReport)

        With dsCommand_TerminateReport
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TerminateReportSet, Table_TerminateReport)

        End With

        TerminateReportSet.AcceptChanges()

    End Function
End Class
