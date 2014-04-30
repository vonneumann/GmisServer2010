Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectMessages

    Public Const Table_Project_Messages As String = "project_task"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectMessages As SqlDataAdapter

    '�����ѯ����
    Private GetWfProjectMessagesInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectMessages = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectMessagesInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfProjectMessagesInfo(ByVal strSQL_Condition_WfProjectMessages As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectMessagesInfoCommand Is Nothing Then

            GetWfProjectMessagesInfoCommand = New SqlCommand("GetWfProjectMessagesInfo", conn)
            GetWfProjectMessagesInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectMessagesInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectMessages
            .SelectCommand = GetWfProjectMessagesInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectMessagesInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectMessages
            .Fill(tempDs, Table_Project_Messages)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateWfProjectMessages(ByVal WfProjectMessagesSet As DataSet)

        If WfProjectMessagesSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfProjectMessagesSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectMessages)

        With dsCommand_WfProjectMessages
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectMessagesSet, Table_Project_Messages)

            WfProjectMessagesSet.AcceptChanges()
        End With
    End Function
End Class
