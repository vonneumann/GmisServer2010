Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTaskTransfer
    Public Const Table_Project_Task_Transfer As String = "project_task_transfer"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectTaskTransfer As SqlDataAdapter

    '�����ѯ����
    Private GetWfProjectTaskTransferInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectTaskTransfer = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectTaskTransferInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfProjectTaskTransferInfo(ByVal strSQL_Condition_WfProjectTaskTransfer As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTaskTransferInfoCommand Is Nothing Then

            GetWfProjectTaskTransferInfoCommand = New SqlCommand("GetWfProjectTaskTransferInfo", conn)
            GetWfProjectTaskTransferInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTaskTransferInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTaskTransfer
            .SelectCommand = GetWfProjectTaskTransferInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTaskTransferInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTaskTransfer
            .Fill(tempDs, Table_Project_Task_Transfer)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateWfProjectTaskTransfer(ByVal WfProjectTaskTransferSet As DataSet)

        If WfProjectTaskTransferSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If WfProjectTaskTransferSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTaskTransfer)

        With dsCommand_WfProjectTaskTransfer
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTaskTransferSet, Table_Project_Task_Transfer)

            WfProjectTaskTransferSet.AcceptChanges()
        End With

    End Function
End Class
