Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTaskTemplate
    Public Const Table_TaskTemplate As String = "task_template"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfTaskTemplate As SqlDataAdapter

    '�����ѯ����
    Private GetWfTaskTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfTaskTemplate = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfTaskTemplateInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfTaskTemplateInfo(ByVal strSQL_Condition_WfTaskTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTaskTemplateInfoCommand Is Nothing Then

            GetWfTaskTemplateInfoCommand = New SqlCommand("GetWfTaskTemplateInfo", conn)
            GetWfTaskTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTaskTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTaskTemplate
            .SelectCommand = GetWfTaskTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTaskTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTaskTemplate
            .Fill(tempDs, Table_TaskTemplate)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTaskTemplate(ByVal WfTaskTemplateSet As DataSet)

        If WfTaskTemplateSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfTaskTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTaskTemplate)

        With dsCommand_WfTaskTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTaskTemplateSet, Table_TaskTemplate)

            WfTaskTemplateSet.AcceptChanges()
        End With

    End Function
End Class
