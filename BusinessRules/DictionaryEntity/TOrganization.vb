Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TOrganization
    Public Const Table_TOrganization As String = "TOrganization"


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_TOrganization As SqlDataAdapter


    '�����ѯ����
    Private GetTOrganizationInfoCommand As SqlCommand


    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_TOrganization = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetTOrganizationInfo("null")
    End Sub


    Public Function GetTOrganizationInfo(ByVal strSQL_Condition_TOrganization) As DataSet

        Dim tempDs As New DataSet()

        If GetTOrganizationInfoCommand Is Nothing Then

            GetTOrganizationInfoCommand = New SqlCommand("GetTOrganizationInfo", conn)
            GetTOrganizationInfoCommand.CommandType = CommandType.StoredProcedure
            GetTOrganizationInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TOrganization
            .SelectCommand = GetTOrganizationInfoCommand
            .SelectCommand.Transaction = ts
            GetTOrganizationInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TOrganization
            .Fill(tempDs, Table_TOrganization)
        End With

        Return tempDs

    End Function


    Public Function UpdateTOrganization(ByVal TOrganization As DataSet)

        If TOrganization Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If TOrganization.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TOrganization)

        With dsCommand_TOrganization
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TOrganization, Table_TOrganization)

            TOrganization.AcceptChanges()
        End With


    End Function

End Class
