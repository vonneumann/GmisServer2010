Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ReturnReceipt

    Public Const Table_ReturnReceipt As String = "loan_return_receipt"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ReturnReceipt As SqlDataAdapter

    '�����ѯ����
    Private GetReturnReceiptInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ReturnReceipt = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetReturnReceiptInfo("null")
    End Sub

    '��ȡ�ſ��ִ��Ϣ
    Public Function GetReturnReceiptInfo(ByVal strSQL_Condition_ReturnReceipt As String) As DataSet

        Dim tempDs As New DataSet()

        If GetReturnReceiptInfoCommand Is Nothing Then

            GetReturnReceiptInfoCommand = New SqlCommand("GetReturnReceiptInfo", conn)
            GetReturnReceiptInfoCommand.CommandType = CommandType.StoredProcedure
            GetReturnReceiptInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ReturnReceipt
            .SelectCommand = GetReturnReceiptInfoCommand
            .SelectCommand.Transaction = ts
            GetReturnReceiptInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ReturnReceipt
            .Fill(tempDs, Table_ReturnReceipt)
        End With

        Return tempDs

    End Function

    '���·ſ��ִ��Ϣ
    Public Function UpdateReturnReceipt(ByVal ReturnReceiptSet As DataSet)

        If ReturnReceiptSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ReturnReceiptSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ReturnReceipt)

        With dsCommand_ReturnReceipt
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ReturnReceiptSet, Table_ReturnReceipt)

        End With

        ReturnReceiptSet.AcceptChanges()


    End Function

End Class
