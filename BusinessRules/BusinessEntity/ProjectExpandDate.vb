Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectExpandDate

    Public Const Table_Project_ProjectExpandDate As String = "project_expand_date"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectExpandDate As SqlDataAdapter

    '�����ѯ����
    Private GetProjectExpandDateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectExpandDate = New SqlDataAdapter

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectExpandDateInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectExpandDateInfo(ByVal strSQL_Condition_ProjectExpandDate As String) As DataSet

        Dim tempDs As New DataSet

        If GetProjectExpandDateInfoCommand Is Nothing Then

            GetProjectExpandDateInfoCommand = New SqlCommand("GetProjectExpandDateInfo", conn)
            GetProjectExpandDateInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectExpandDateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectExpandDate
            .SelectCommand = GetProjectExpandDateInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectExpandDateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectExpandDate
            .Fill(tempDs, Table_Project_ProjectExpandDate)
        End With

        Return tempDs

    End Function


    '������Ŀ������Ϣ
    Public Function UpdateProjectExpandDate(ByVal ProjectExpandDateSet As DataSet)

        If ProjectExpandDateSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectExpandDateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectExpandDate)

        With dsCommand_ProjectExpandDate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectExpandDateSet, Table_Project_ProjectExpandDate)

            ProjectExpandDateSet.AcceptChanges()
        End With


    End Function


End Class

