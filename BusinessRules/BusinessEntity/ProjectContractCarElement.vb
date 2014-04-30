Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectContractCarElement

    Public Const Table_ProjectContractCarElement As String = "project_contract_car_element"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectContractCarElement As SqlDataAdapter

    '�����ѯ����
    Private GetProjectContractCarElementInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectContractCarElement = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectContractCarElementInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectContractCarElementInfo(ByVal strSQL_Condition_ProjectContractCarElement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectContractCarElementInfoCommand Is Nothing Then

            GetProjectContractCarElementInfoCommand = New SqlCommand("GetProjectContractCarElementInfo", conn)
            GetProjectContractCarElementInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectContractCarElementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectContractCarElement
            .SelectCommand = GetProjectContractCarElementInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectContractCarElementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectContractCarElement
            .Fill(tempDs, Table_ProjectContractCarElement)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProjectContractCarElement(ByVal ProjectContractCarElementSet As DataSet)

        If ProjectContractCarElementSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectContractCarElementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectContractCarElement)

        With dsCommand_ProjectContractCarElement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectContractCarElementSet, Table_ProjectContractCarElement)

        End With

        ProjectContractCarElementSet.AcceptChanges()

    End Function
End Class
