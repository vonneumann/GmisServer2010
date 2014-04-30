Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class RefundCertificate

    Public Const Table_RefundCertificate As String = "refund_certificate"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_RefundCertificate As SqlDataAdapter

    '�����ѯ����
    Private GetRefundCertificateInfoCommand As SqlCommand
    Private GetMaxRefundCertificateNumCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_RefundCertificate = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetRefundCertificateInfo("null")
    End Sub

    '��ȡ����֤������Ϣ
    Public Function GetRefundCertificateInfo(ByVal strSQL_Condition_RefundCertificate As String) As DataSet

        Dim tempDs As New DataSet()

        If GetRefundCertificateInfoCommand Is Nothing Then

            GetRefundCertificateInfoCommand = New SqlCommand("GetRefundCertificateInfo", conn)
            GetRefundCertificateInfoCommand.CommandType = CommandType.StoredProcedure
            GetRefundCertificateInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_RefundCertificate
            .SelectCommand = GetRefundCertificateInfoCommand
            .SelectCommand.Transaction = ts
            GetRefundCertificateInfoCommand.Parameters("@Condition").Value = strSQL_Condition_RefundCertificate
            .Fill(tempDs, Table_RefundCertificate)
        End With

        Return tempDs

    End Function

    '���»���֤������Ϣ
    Public Function UpdateRefundCertificate(ByVal RefundCertificateSet As DataSet)

        If RefundCertificateSet Is Nothing Then
            Exit Function
        End If



        '�����¼��δ�����κα仯�����˳�����
        If RefundCertificateSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_RefundCertificate)

        With dsCommand_RefundCertificate
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(RefundCertificateSet, Table_RefundCertificate)

        End With

        RefundCertificateSet.AcceptChanges()

    End Function
End Class
