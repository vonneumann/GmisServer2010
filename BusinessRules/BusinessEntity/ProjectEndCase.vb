Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectEndCase

    Public Const Table_ProjectEndCase As String = "project_end_case"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectEndCase As SqlDataAdapter

    '�����ѯ����
    Private GetProjectEndCaseInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectEndCase = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectEndCaseInfo("null")

    End Sub

    '��ȡ������¼��Ϣ
    Public Function GetProjectEndCaseInfo(ByVal strSQL_Condition_WorkLog As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectEndCaseInfoCommand Is Nothing Then

            GetProjectEndCaseInfoCommand = New SqlCommand("GetProjectEndCaseInfo", conn)
            GetProjectEndCaseInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectEndCaseInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectEndCase
            .SelectCommand = GetProjectEndCaseInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectEndCaseInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WorkLog
            .Fill(tempDs, Table_ProjectEndCase)
        End With

        Return tempDs


    End Function

    '���¹�����¼��Ϣ
    Public Function UpdateProjectEndCase(ByVal ProjectEndCase As DataSet)

        If ProjectEndCase Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectEndCase.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectEndCase)

        With dsCommand_ProjectEndCase
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectEndCase, Table_ProjectEndCase)

            ProjectEndCase.AcceptChanges()
        End With


    End Function

End Class
