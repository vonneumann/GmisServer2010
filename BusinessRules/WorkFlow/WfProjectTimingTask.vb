Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTimingTask

    Public Const Table_Project_Timing_Task As String = "project_timing_task"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectTimingTask As SqlDataAdapter

    '�����ѯ����
    Private GetWfProjectTimingTaskInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectTimingTask = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectTimingTaskInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfProjectTimingTaskInfo(ByVal strSQL_Condition_WfProjectTimingTask As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTimingTaskInfoCommand Is Nothing Then

            GetWfProjectTimingTaskInfoCommand = New SqlCommand("GetWfProjectTimingTaskInfo", conn)
            GetWfProjectTimingTaskInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTimingTaskInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTimingTask
            .SelectCommand = GetWfProjectTimingTaskInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTimingTaskInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTimingTask
            .Fill(tempDs, Table_Project_Timing_Task)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateWfProjectTimingTask(ByVal WfProjectTimingTaskSet As DataSet)

        If WfProjectTimingTaskSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If WfProjectTimingTaskSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTimingTask)

        With dsCommand_WfProjectTimingTask
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTimingTaskSet, Table_Project_Timing_Task)

            WfProjectTimingTaskSet.AcceptChanges()
        End With


    End Function
End Class
