
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTaskTransferTemplate
    Public Const Table_Task_Transfer_Template As String = "task_transfer_template"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfTaskTransferTemplate As SqlDataAdapter

    '�����ѯ����
    Private GetWfTaskTransferTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfTaskTransferTemplate = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfTaskTransferTemplateInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfTaskTransferTemplateInfo(ByVal strSQL_Condition_WfTaskTransferTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTaskTransferTemplateInfoCommand Is Nothing Then

            GetWfTaskTransferTemplateInfoCommand = New SqlCommand("GetWfTaskTransferTemplateInfo", conn)
            GetWfTaskTransferTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTaskTransferTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTaskTransferTemplate
            .SelectCommand = GetWfTaskTransferTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTaskTransferTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTaskTransferTemplate
            .Fill(tempDs, Table_Task_Transfer_Template)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTaskTransferTemplate(ByVal WfTaskTransferTemplateSet As DataSet)

        If WfTaskTransferTemplateSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfTaskTransferTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTaskTransferTemplate)

        With dsCommand_WfTaskTransferTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTaskTransferTemplateSet, Table_Task_Transfer_Template)

            WfTaskTransferTemplateSet.AcceptChanges()
        End With

    End Function
End Class
