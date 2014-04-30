Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class JobPeriod
    Private Const Table_JobPeriod As String = "TJobPeriod"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_JobPeriod As SqlDataAdapter

    '�����ѯ����
    Private GetJobPeriodCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_JobPeriod = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetJobPeriodInfo("null")
    End Sub

    '��ȡ����ʱ�����Ϣ
    Public Function GetJobPeriodInfo(ByVal strSQL_Condition_JobPeriod As String) As DataSet

        Dim tempDs As New DataSet()

        If GetJobPeriodCommand Is Nothing Then

            GetJobPeriodCommand = New SqlCommand("dbo.GetJobPeriodInfo", conn)
            GetJobPeriodCommand.CommandType = CommandType.StoredProcedure
            GetJobPeriodCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_JobPeriod
            .SelectCommand = GetJobPeriodCommand
            .SelectCommand.Transaction = ts
            GetJobPeriodCommand.Parameters("@Condition").Value = strSQL_Condition_JobPeriod
            .Fill(tempDs, Table_JobPeriod)
        End With

        Return tempDs

    End Function

    '���¼�����Ϣ
    Public Function UpdateJobPeriod(ByVal JobPeriodSet As DataSet)

        If JobPeriodSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If JobPeriodSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_JobPeriod)

        With dsCommand_JobPeriod
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(JobPeriodSet, Table_JobPeriod)

        End With

        JobPeriodSet.AcceptChanges()
    End Function
End Class
