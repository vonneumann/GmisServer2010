Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectSignature

    Public Const Table_Project_Signature As String = "project_signature"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectSignature As SqlDataAdapter

    '�����ѯ����
    Private GetProjectSignatureInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectSignature = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectSignatureInfo("null")
    End Sub

    '��ȡǩԼ��Ϣ
    Public Function GetProjectSignatureInfo(ByVal strSQL_Condition_ProjectSignature As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectSignatureInfoCommand Is Nothing Then

            GetProjectSignatureInfoCommand = New SqlCommand("GetProjectSignatureInfo", conn)
            GetProjectSignatureInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectSignatureInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectSignature
            .SelectCommand = GetProjectSignatureInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectSignatureInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectSignature
            .Fill(tempDs, Table_Project_Signature)
        End With

        Return tempDs

    End Function


    '����ǩԼ��Ϣ
    Public Function UpdateProjectSignature(ByVal ProjectSignatureSet As DataSet)


        If ProjectSignatureSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectSignatureSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectSignature)

        With dsCommand_ProjectSignature
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectSignatureSet, Table_Project_Signature)

            ProjectSignatureSet.AcceptChanges()
        End With


    End Function

End Class
