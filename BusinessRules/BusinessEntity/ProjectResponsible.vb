Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectResponsible

    Public Const Table_ProjectResponsible As String = "project_responsible"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectResponsible As SqlDataAdapter

    '�����ѯ����
    Private GetProjectResponsibleInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectResponsible = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectResponsibleInfo("null")

    End Sub

    '��ȡ������¼��Ϣ
    Public Function GetProjectResponsibleInfo(ByVal strSQL_Condition As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectResponsibleInfoCommand Is Nothing Then

            GetProjectResponsibleInfoCommand = New SqlCommand("GetProjectResponsible", conn)
            GetProjectResponsibleInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectResponsibleInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectResponsible
            .SelectCommand = GetProjectResponsibleInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectResponsibleInfoCommand.Parameters("@Condition").Value = strSQL_Condition
            .Fill(tempDs, Table_ProjectResponsible)
        End With

        Return tempDs


    End Function

    '���¹�����¼��Ϣ
    Public Function UpdateProjectResponsible(ByVal ProjectResponsible As DataSet)

        If ProjectResponsible Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectResponsible.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectResponsible)

        With dsCommand_ProjectResponsible
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectResponsible, Table_ProjectResponsible)

            ProjectResponsible.AcceptChanges()
        End With


    End Function

End Class
