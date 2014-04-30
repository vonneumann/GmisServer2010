Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTask

    Public Const Table_Project_Task As String = "project_task"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectTask As SqlDataAdapter

    '�����ѯ����
    Private GetWfTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectTask = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectTaskInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfProjectTaskInfo(ByVal strSQL_Condition_WfProjectTask As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTemplateInfoCommand Is Nothing Then

            GetWfTemplateInfoCommand = New SqlCommand("GetWfProjectTaskInfo", conn)
            GetWfTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTask
            .SelectCommand = GetWfTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTask
            .Fill(tempDs, Table_Project_Task)
        End With
        Return tempDs
      
    End Function

    '������Ŀ������Ϣ
    Public Function UpdateWfProjectTask(ByVal WfProjectTaskSet As DataSet)

        If WfProjectTaskSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfProjectTaskSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTask)

        With dsCommand_WfProjectTask
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTaskSet, Table_Project_Task)

            WfProjectTaskSet.AcceptChanges()
        End With


    End Function



End Class
