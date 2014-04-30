Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTimingTaskTemplate
    Public Const Table_Timing_Task_Template As String = "timing_task_template"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfTimingTaskTemplate As SqlDataAdapter

    '�����ѯ����
    Private GetWfTimingTaskTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfTimingTaskTemplate = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfTimingTaskTemplateInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfTimingTaskTemplateInfo(ByVal strSQL_Condition_WfTimingTaskTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTimingTaskTemplateInfoCommand Is Nothing Then

            GetWfTimingTaskTemplateInfoCommand = New SqlCommand("GetWfTimingTaskTemplateInfo", conn)
            GetWfTimingTaskTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTimingTaskTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTimingTaskTemplate
            .SelectCommand = GetWfTimingTaskTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTimingTaskTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTimingTaskTemplate
            .Fill(tempDs, Table_Timing_Task_Template)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTimingTaskTemplate(ByVal WfTimingTaskTemplateSet As DataSet)

        If WfTimingTaskTemplateSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfTimingTaskTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTimingTaskTemplate)

        With dsCommand_WfTimingTaskTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTimingTaskTemplateSet, Table_Timing_Task_Template)

            WfTimingTaskTemplateSet.AcceptChanges()
        End With


    End Function

End Class
