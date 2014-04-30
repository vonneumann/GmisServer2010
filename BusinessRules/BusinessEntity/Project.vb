Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Project

    Public Const Table_Project As String = "project"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Project As SqlDataAdapter

    '�����ѯ����
    Private GetProjectInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Project = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectInfo(ByVal strSQL_Condition_Project As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectInfoCommand Is Nothing Then

            GetProjectInfoCommand = New SqlCommand("GetProjectInfo", conn)
            GetProjectInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Project
            .SelectCommand = GetProjectInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Project
            .Fill(tempDs, Table_Project)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProject(ByVal ProjectSet As DataSet)

        If ProjectSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Project)

        With dsCommand_Project
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectSet, Table_Project)

        End With

        ProjectSet.AcceptChanges()

    End Function
End Class
