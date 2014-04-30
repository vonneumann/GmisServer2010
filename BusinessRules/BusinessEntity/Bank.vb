Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Bank


    Public Const Table_Bank As String = "bank"
    Public Const Table_Bank_Branch As String = "bank_branch"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Bank As SqlDataAdapter
    Private dsCommand_Branch As SqlDataAdapter

    '�����ѯ����
    Private GetBankInfoCommand As SqlCommand
    Private GetBranchInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Bank = New SqlDataAdapter()
        dsCommand_Branch = New SqlDataAdapter()


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetBankInfo("null", "null")
    End Sub

    '��ȡ������Ϣ
    Public Function GetBankInfo(ByVal strSQL_Condition_Bank As String, ByVal strSQL_Condition_Branch As String) As DataSet

        Dim tempDs As New DataSet()

        If GetBankInfoCommand Is Nothing Then

            GetBankInfoCommand = New SqlCommand("GetBankInfo", conn)
            GetBankInfoCommand.CommandType = CommandType.StoredProcedure
            GetBankInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Bank
            .SelectCommand = GetBankInfoCommand
            .SelectCommand.Transaction = ts
            GetBankInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Bank
            .Fill(tempDs, Table_Bank)
        End With

        If GetBranchInfoCommand Is Nothing Then

            GetBranchInfoCommand = New SqlCommand("GetBankBranchInfo", conn)
            GetBranchInfoCommand.CommandType = CommandType.StoredProcedure
            GetBranchInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Branch
            .SelectCommand = GetBranchInfoCommand
            .SelectCommand.Transaction = ts
            GetBranchInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Branch
            .Fill(tempDs, Table_Bank_Branch)
        End With



        GetBankInfo = tempDs

    End Function

    '����������Ϣ
    Public Function UpdateBank(ByVal BankSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Bank)

        With dsCommand_Bank

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(BankSet, Table_Bank)

        End With

    End Function

    '����֧����Ϣ
    Public Function UpdateBranch(ByVal BranchSet As DataSet)


        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Branch)

        With dsCommand_Branch

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(BranchSet, Table_Bank_Branch)

        End With



    End Function

    '��������֧����Ϣ
    Public Function UpdateBankAndBranch(ByVal BankAndBranchSet As DataSet)

        If BankAndBranchSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If BankAndBranchSet.HasChanges = False Then
            Exit Function
        End If


        'ɾ������
        If IsNothing(BankAndBranchSet.GetChanges(DataRowState.Deleted)) = False Then
            '��ɾ��ϸ����ɾ����
            UpdateBranch(BankAndBranchSet.GetChanges(DataRowState.Deleted))
            UpdateBank(BankAndBranchSet.GetChanges(DataRowState.Deleted))

        End If

        '������
        If IsNothing(BankAndBranchSet.GetChanges(DataRowState.Added)) = False Then
            UpdateBank(BankAndBranchSet.GetChanges(DataRowState.Added))
            UpdateBranch(BankAndBranchSet.GetChanges(DataRowState.Added))
        End If

        '���²���
        If IsNothing(BankAndBranchSet.GetChanges(DataRowState.Modified)) = False Then
            UpdateBank(BankAndBranchSet.GetChanges(DataRowState.Modified))
            UpdateBranch(BankAndBranchSet.GetChanges(DataRowState.Modified))
        End If

        BankAndBranchSet.AcceptChanges()
    End Function


End Class
