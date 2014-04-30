Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class LoanNotice
    Public Const Table_LoanNotice As String = "loan_notice"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_LoanNotice As SqlDataAdapter

    '�����ѯ����
    Private GetLoanNoticeInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_LoanNotice = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetLoanNoticeInfo("null")
    End Sub

    '��ȡ�ſ�֪ͨ����Ϣ
    Public Function GetLoanNoticeInfo(ByVal strSQL_Condition_LoanNotice As String) As DataSet

        Dim tempDs As New DataSet()

        If GetLoanNoticeInfoCommand Is Nothing Then

            GetLoanNoticeInfoCommand = New SqlCommand("GetLoanNoticeInfo", conn)
            GetLoanNoticeInfoCommand.CommandType = CommandType.StoredProcedure
            GetLoanNoticeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_LoanNotice
            .SelectCommand = GetLoanNoticeInfoCommand
            .SelectCommand.Transaction = ts
            GetLoanNoticeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_LoanNotice
            .Fill(tempDs, Table_LoanNotice)
        End With

        Return tempDs

    End Function

    '���·ſ�֪ͨ����Ϣ
    Public Function UpdateLoanNotice(ByVal LoanNoticeSet As DataSet)

        If LoanNoticeSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If LoanNoticeSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_LoanNotice)

        With dsCommand_LoanNotice
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(LoanNoticeSet, Table_LoanNotice)

        End With

        LoanNoticeSet.AcceptChanges()

    End Function
End Class
