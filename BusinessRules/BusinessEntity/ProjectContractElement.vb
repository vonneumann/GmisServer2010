Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractElement

    Public Const Table_ProjectContractElement As String = "project_contract_element"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectContractElement As SqlDataAdapter

    '�����ѯ����
    Private GetProjectContractElementInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectContractElement = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectContractElementInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectContractElementInfo(ByVal strSQL_Condition_ProjectContractElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractElementInfoCommand Is Nothing Then

            GetProjectContractElementInfoCommand = New SqlCommand("GetProjectContractElementInfo", conn)
            GetProjectContractElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractElement
            .SelectCommand = GetProjectContractElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractElement
            .Fill(tempDs, Table_ProjectContractElement)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProjectContractElement(ByVal ProjectContractElementSet As DataSet)

        If ProjectContractElementSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectContractElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractElement)

        With dsCommand_ProjectContractElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractElementSet, Table_ProjectContractElement)

        End With

        ProjectContractElementSet.AcceptChanges()

    End Function
End Class
