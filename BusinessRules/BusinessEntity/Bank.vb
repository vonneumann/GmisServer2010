Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Bank


    Public Const Table_Bank As String = "bank"
    Public Const Table_Bank_Branch As String = "bank_branch"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Bank As SqlDataAdapter
    Private dsCommand_Branch As SqlDataAdapter

    '定义查询命令
    Private GetBankInfoCommand As SqlCommand
    Private GetBranchInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Bank = New SqlDataAdapter()
        dsCommand_Branch = New SqlDataAdapter()


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetBankInfo("null", "null")
    End Sub

    '获取银行信息
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

    '更新银行信息
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

    '更新支行信息
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

    '更新银行支行信息
    Public Function UpdateBankAndBranch(ByVal BankAndBranchSet As DataSet)

        If BankAndBranchSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If BankAndBranchSet.HasChanges = False Then
            Exit Function
        End If


        '删除操作
        If IsNothing(BankAndBranchSet.GetChanges(DataRowState.Deleted)) = False Then
            '先删明细表，再删主表
            UpdateBranch(BankAndBranchSet.GetChanges(DataRowState.Deleted))
            UpdateBank(BankAndBranchSet.GetChanges(DataRowState.Deleted))

        End If

        '新增作
        If IsNothing(BankAndBranchSet.GetChanges(DataRowState.Added)) = False Then
            UpdateBank(BankAndBranchSet.GetChanges(DataRowState.Added))
            UpdateBranch(BankAndBranchSet.GetChanges(DataRowState.Added))
        End If

        '更新操作
        If IsNothing(BankAndBranchSet.GetChanges(DataRowState.Modified)) = False Then
            UpdateBank(BankAndBranchSet.GetChanges(DataRowState.Modified))
            UpdateBranch(BankAndBranchSet.GetChanges(DataRowState.Modified))
        End If

        BankAndBranchSet.AcceptChanges()
    End Function


End Class
