
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfTaskRoleTemplate
    Public Const Table_Task_Role_Template As String = "task_role_template"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfTaskRoleTemplate As SqlDataAdapter

    '�����ѯ����
    Private GetWfTaskRoleTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfTaskRoleTemplate = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfTaskRoleTemplateInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfTaskRoleTemplateInfo(ByVal strSQL_Condition_WfTaskRoleTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTaskRoleTemplateInfoCommand Is Nothing Then

            GetWfTaskRoleTemplateInfoCommand = New SqlCommand("GetWfTaskRoleTemplateInfo", conn)
            GetWfTaskRoleTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTaskRoleTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfTaskRoleTemplate
            .SelectCommand = GetWfTaskRoleTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTaskRoleTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfTaskRoleTemplate
            .Fill(tempDs, Table_Task_Role_Template)
        End With

        Return tempDs

    End Function

    Public Function UpdateWfTaskRoleTemplate(ByVal WfTaskRoleTemplateSet As DataSet)

        If WfTaskRoleTemplateSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfTaskRoleTemplateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfTaskRoleTemplate)

        With dsCommand_WfTaskRoleTemplate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfTaskRoleTemplateSet, Table_Task_Role_Template)

            WfTaskRoleTemplateSet.AcceptChanges()
        End With


    End Function
End Class
