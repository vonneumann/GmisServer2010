Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Process

    Public Const Table_Process As String = "process"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Process As SqlDataAdapter

    '�����ѯ����
    Private GetProcessInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Process = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProcessInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProcessInfo(ByVal strSQL_Condition_Process As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProcessInfoCommand Is Nothing Then

            GetProcessInfoCommand = New SqlCommand("GetProcessInfo", conn)
            GetProcessInfoCommand.CommandType = CommandType.StoredProcedure
            GetProcessInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Process
            .SelectCommand = GetProcessInfoCommand
            .SelectCommand.Transaction = ts
            GetProcessInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Process
            .Fill(tempDs, Table_Process)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProcess(ByVal ProcessSet As DataSet)

        If ProcessSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If ProcessSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Process)

        With dsCommand_Process
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProcessSet, Table_Process)

        End With

        ProcessSet.AcceptChanges()
    End Function
End Class
