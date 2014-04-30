Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TimeServerLog
    Public Const Table_Project_Task_Attendee As String = "time_server_log"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_TimeServerLog As SqlDataAdapter

    '�����ѯ����
    Private GetTimeServerLogInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_TimeServerLog = New SqlDataAdapter

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetTimeServerLogInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetTimeServerLogInfo(ByVal strSQL_Condition_TimeServerLog As String) As DataSet

        Dim tempDs As New DataSet

        If GetTimeServerLogInfoCommand Is Nothing Then

            GetTimeServerLogInfoCommand = New SqlCommand("GetTimeServerLogInfo", conn)
            GetTimeServerLogInfoCommand.CommandType = CommandType.StoredProcedure
            GetTimeServerLogInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TimeServerLog
            .SelectCommand = GetTimeServerLogInfoCommand
            .SelectCommand.Transaction = ts
            GetTimeServerLogInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TimeServerLog
            .Fill(tempDs, Table_Project_Task_Attendee)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateTimeServerLog(ByVal TimeServerLogSet As DataSet)


        If TimeServerLogSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If TimeServerLogSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TimeServerLog)

        With dsCommand_TimeServerLog
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TimeServerLogSet, Table_Project_Task_Attendee)

            TimeServerLogSet.AcceptChanges()

        End With


    End Function

End Class
