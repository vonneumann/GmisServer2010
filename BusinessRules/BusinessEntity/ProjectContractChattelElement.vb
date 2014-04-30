Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractChattelElement

    Public Const Table_ProjectContractChattelElement As String = "project_contract_chattel_element"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectContractChattelElement As SqlDataAdapter

    '�����ѯ����
    Private GetProjectContractChattelElementInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectContractChattelElement = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectContractChattelElementInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectContractChattelElementInfo(ByVal strSQL_Condition_ProjectContractChattelElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractChattelElementInfoCommand Is Nothing Then

            GetProjectContractChattelElementInfoCommand = New SqlCommand("GetProjectContractChattelElementInfo", conn)
            GetProjectContractChattelElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractChattelElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractChattelElement
            .SelectCommand = GetProjectContractChattelElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractChattelElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractChattelElement
            .Fill(tempDs, Table_ProjectContractChattelElement)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProjectContractChattelElement(ByVal ProjectContractChattelElementSet As DataSet)

        If ProjectContractChattelElementSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectContractChattelElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractChattelElement)

        With dsCommand_ProjectContractChattelElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractChattelElementSet, Table_ProjectContractChattelElement)

        End With

        ProjectContractChattelElementSet.AcceptChanges()

    End Function
End Class
