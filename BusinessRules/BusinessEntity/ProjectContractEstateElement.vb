Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractEstateElement

    Public Const Table_ProjectContractEstateElement As String = "project_contract_estate_element"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectContractEstateElement As SqlDataAdapter

    '�����ѯ����
    Private GetProjectContractEstateElementInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectContractEstateElement = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectContractEstateElementInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectContractEstateElementInfo(ByVal strSQL_Condition_ProjectContractEstateElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractEstateElementInfoCommand Is Nothing Then

            GetProjectContractEstateElementInfoCommand = New SqlCommand("GetProjectContractEstateElementInfo", conn)
            GetProjectContractEstateElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractEstateElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractEstateElement
            .SelectCommand = GetProjectContractEstateElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractEstateElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractEstateElement
            .Fill(tempDs, Table_ProjectContractEstateElement)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProjectContractEstateElement(ByVal ProjectContractEstateElementSet As DataSet)

        If ProjectContractEstateElementSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectContractEstateElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractEstateElement)

        With dsCommand_ProjectContractEstateElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractEstateElementSet, Table_ProjectContractEstateElement)

        End With

        ProjectContractEstateElementSet.AcceptChanges()

    End Function
End Class
