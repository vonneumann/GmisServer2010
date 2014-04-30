Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectPrompt
    Public Const Table_Project_Prompt As String = "project_prompt"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectPrompt As SqlDataAdapter

    '�����ѯ����
    Private GetWfTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectPrompt = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectPromptInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfProjectPromptInfo(ByVal strSQL_Condition_WfProjectPrompt As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfTemplateInfoCommand Is Nothing Then

            GetWfTemplateInfoCommand = New SqlCommand("GetWfProjectPromptInfo", conn)
            GetWfTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectPrompt
            .SelectCommand = GetWfTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectPrompt
            .Fill(tempDs, Table_Project_Prompt)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateWfProjectPrompt(ByVal WfProjectPromptSet As DataSet)


        If WfProjectPromptSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfProjectPromptSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectPrompt)

        With dsCommand_WfProjectPrompt
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectPromptSet, Table_Project_Prompt)

            WfProjectPromptSet.AcceptChanges()
        End With


    End Function
End Class
