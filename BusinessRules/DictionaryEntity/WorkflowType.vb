Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WorkflowType
    Public Const Table_WorkflowType As String = "workflow"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WorkflowType As SqlDataAdapter

    '�����ѯ����
    Private GetWorkflowTypeInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WorkflowType = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWorkflowTypeInfo("null")
    End Sub

    '��ȡ������������Ϣ
    Public Function GetWorkflowTypeInfo(ByVal strSQL_Condition_WorkflowType As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWorkflowTypeInfoCommand Is Nothing Then

            GetWorkflowTypeInfoCommand = New SqlCommand("GetWorkflowTypeInfo", conn)
            GetWorkflowTypeInfoCommand.CommandType = CommandType.StoredProcedure
            GetWorkflowTypeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WorkflowType
            .SelectCommand = GetWorkflowTypeInfoCommand
            .SelectCommand.Transaction = ts
            GetWorkflowTypeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WorkflowType
            .Fill(tempDs, Table_WorkflowType)
        End With

        Return tempDs


    End Function

    '���¹�����������Ϣ
    Public Function UpdateWorkflowType(ByVal WorkflowTypeSet As DataSet)

        '�����¼��δ�����κα仯�����˳�����
        If WorkflowTypeSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WorkflowType)

        With dsCommand_WorkflowType
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WorkflowTypeSet, Table_WorkflowType)

        End With

        WorkflowTypeSet.AcceptChanges()
    End Function
End Class
