
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfMessagesTemplate
    Public Const Table_Messages_Dict As String = "messages_dict"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfMessagesTemplate As SqlDataAdapter

    '�����ѯ����
    Private GetWfMessagesTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfMessagesTemplate = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfMessagesTemplateInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfMessagesTemplateInfo(ByVal strSQL_Condition_WfMessagesTemplate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfMessagesTemplateInfoCommand Is Nothing Then

            GetWfMessagesTemplateInfoCommand = New SqlCommand("GetWfMessagesTemplateInfo", conn)
            GetWfMessagesTemplateInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfMessagesTemplateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfMessagesTemplate
            .SelectCommand = GetWfMessagesTemplateInfoCommand
            .SelectCommand.Transaction = ts
            GetWfMessagesTemplateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfMessagesTemplate
            .Fill(tempDs, Table_Messages_Dict)
        End With

        Return tempDs

    End Function

End Class
