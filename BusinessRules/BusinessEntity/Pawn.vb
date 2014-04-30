Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Pawn

    'Public Const Table_Pawn_Credit As String = "TPawn_Credit"
    Public Const Table_Pawn As String = "TPawn"
    Public Const Table_Pawn_Continue As String = "TPawn_Continue"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    'Private dsCommand_Pawn_Credit As SqlDataAdapter
    Private dsCommand_Pawn As SqlDataAdapter
    Private dsCommand_Pawn_Continue As SqlDataAdapter

    '定义查询命令
    'Private GetPawnCreditInfoCommand As SqlCommand
    Private GetPawnInfoCommand As SqlCommand
    Private GetPawnContinueInfoCommand As SqlCommand
    Private GetMaxPawnNumCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        'dsCommand_Pawn_Credit = New SqlDataAdapter
        dsCommand_Pawn = New SqlDataAdapter
        dsCommand_Pawn_Continue = New SqlDataAdapter

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetPawnInfo("null", "null")
    End Sub

    '获取项目评价信息
    Public Function GetPawnInfo(ByVal strSQL_Condition_Pawn As String, ByVal strSQL_Condition_Pawn_Continue As String) As DataSet

        Dim tempDs As New DataSet

        'If GetPawnCreditInfoCommand Is Nothing Then

        '    GetPawnCreditInfoCommand = New SqlCommand("GetPawnCreditInfo", conn)
        '    GetPawnCreditInfoCommand.CommandType = CommandType.StoredProcedure
        '    GetPawnCreditInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        'End If

        'With dsCommand_Pawn_Credit
        '    .SelectCommand = GetPawnCreditInfoCommand
        '    .SelectCommand.Transaction = ts
        '    GetPawnCreditInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Pawn_Credit
        '    .Fill(tempDs, Table_Pawn_Credit)
        'End With

        If GetPawnInfoCommand Is Nothing Then

            GetPawnInfoCommand = New SqlCommand("GetPawnInfo", conn)
            GetPawnInfoCommand.CommandType = CommandType.StoredProcedure
            GetPawnInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Pawn
            .SelectCommand = GetPawnInfoCommand
            .SelectCommand.Transaction = ts
            GetPawnInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Pawn
            .Fill(tempDs, Table_Pawn)
        End With

        If GetPawnContinueInfoCommand Is Nothing Then

            GetPawnContinueInfoCommand = New SqlCommand("GetPawnContinueInfo", conn)
            GetPawnContinueInfoCommand.CommandType = CommandType.StoredProcedure
            GetPawnContinueInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Pawn_Continue
            .SelectCommand = GetPawnContinueInfoCommand
            .SelectCommand.Transaction = ts
            GetPawnContinueInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Pawn_Continue
            .Fill(tempDs, Table_Pawn_Continue)
        End With

        'tempDs.Relations.Add("TPawn_TPawnContinue", tempDs.Tables(Table_Pawn).Columns("pawn_project_code"), tempDs.Tables(Table_Pawn_Continue).Columns("pawn_project_code"))

        Return tempDs

    End Function

    '获取最大序列号
    Public Function GetMaxPawnNum(ByVal projectID As String) As Integer

        If GetMaxPawnNumCommand Is Nothing Then

            GetMaxPawnNumCommand = New SqlCommand("GetMaxPawnNum", conn)
            GetMaxPawnNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxPawnNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxPawnNumCommand.Parameters.Add(New SqlParameter("@maxPawnNum", SqlDbType.Int))
            GetMaxPawnNumCommand.Parameters.Item("@maxPawnNum").Direction = ParameterDirection.Output
            GetMaxPawnNumCommand.Transaction = ts
        End If

        GetMaxPawnNumCommand.Parameters("@projectID").Value = projectID
        GetMaxPawnNumCommand.ExecuteNonQuery()
        GetMaxPawnNum = GetMaxPawnNumCommand.Parameters.Item("@maxPawnNum").Value
    End Function

    ''更新项目评价信息
    'Public Function UpdatePawnCredit(ByVal PawnCreditSet As DataSet)

    '    If PawnCreditSet Is Nothing Then
    '        Exit Function
    '    End If

    '    '如果记录集未发生任何变化，则退出过程
    '    If PawnCreditSet.HasChanges = False Then
    '        Exit Function
    '    End If

    '    Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Pawn_Credit)

    '    With dsCommand_Pawn_Credit
    '        .InsertCommand = bd.GetInsertCommand
    '        .UpdateCommand = bd.GetUpdateCommand
    '        .DeleteCommand = bd.GetDeleteCommand

    '        .InsertCommand.Transaction = ts
    '        .UpdateCommand.Transaction = ts
    '        .DeleteCommand.Transaction = ts

    '        .Update(PawnCreditSet, Table_Pawn_Credit)

    '        PawnCreditSet.AcceptChanges()
    '    End With


    'End Function

    Public Function UpdatePawn(ByVal PawnSet As DataSet)

        If PawnSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If PawnSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Pawn)

        With dsCommand_Pawn
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(PawnSet, Table_Pawn)

            PawnSet.AcceptChanges()
        End With


    End Function

    Public Function UpdatePawnContinue(ByVal PawnContinueSet As DataSet)

        If PawnContinueSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If PawnContinueSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Pawn_Continue)

        With dsCommand_Pawn_Continue
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(PawnContinueSet, Table_Pawn_Continue)

            PawnContinueSet.AcceptChanges()
        End With

    End Function

    '更新银行支行信息
    Public Function UpdatePawnAndPawnContinue(ByVal PawnAndPawnContinueSet As DataSet)

        If PawnAndPawnContinueSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If PawnAndPawnContinueSet.HasChanges = False Then
            Exit Function
        End If


        '删除操作
        If IsNothing(PawnAndPawnContinueSet.GetChanges(DataRowState.Deleted)) = False Then
            '先删明细表，再删主表
            UpdatePawnContinue(PawnAndPawnContinueSet.GetChanges(DataRowState.Deleted))
            UpdatePawn(PawnAndPawnContinueSet.GetChanges(DataRowState.Deleted))
            'UpdatePawnCredit(PawnCreditAndPawnAndPawnContinueSet.GetChanges(DataRowState.Deleted))

        End If

        '新增作
        If IsNothing(PawnAndPawnContinueSet.GetChanges(DataRowState.Added)) = False Then
            'UpdatePawnCredit(PawnCreditAndPawnAndPawnContinueSet.GetChanges(DataRowState.Added))
            UpdatePawn(PawnAndPawnContinueSet.GetChanges(DataRowState.Added))
            UpdatePawnContinue(PawnAndPawnContinueSet.GetChanges(DataRowState.Added))
        End If

        '更新操作
        If IsNothing(PawnAndPawnContinueSet.GetChanges(DataRowState.Modified)) = False Then
            'UpdatePawnCredit(PawnCreditAndPawnAndPawnContinueSet.GetChanges(DataRowState.Modified))
            UpdatePawn(PawnAndPawnContinueSet.GetChanges(DataRowState.Modified))
            UpdatePawnContinue(PawnAndPawnContinueSet.GetChanges(DataRowState.Modified))
        End If

        PawnAndPawnContinueSet.AcceptChanges()
    End Function


End Class

