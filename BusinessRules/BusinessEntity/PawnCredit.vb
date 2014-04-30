Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class PawnCredit

    Public Const Table_Pawn_Credit As String = "TPawn_Credit"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Pawn_Credit As SqlDataAdapter


    '定义查询命令
    Private GetPawnCreditInfoCommand As SqlCommand

    Private GetMaxPawnNumCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Pawn_Credit = New SqlDataAdapter

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetPawnCreditInfo("null")
    End Sub

    '获取项目评价信息
    Public Function GetPawnCreditInfo(ByVal strSQL_Condition_Pawn_Credit As String) As DataSet

        Dim tempDs As New DataSet

        If GetPawnCreditInfoCommand Is Nothing Then

            GetPawnCreditInfoCommand = New SqlCommand("GetPawnCreditInfo", conn)
            GetPawnCreditInfoCommand.CommandType = CommandType.StoredProcedure
            GetPawnCreditInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Pawn_Credit
            .SelectCommand = GetPawnCreditInfoCommand
            .SelectCommand.Transaction = ts
            GetPawnCreditInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Pawn_Credit
            .Fill(tempDs, Table_Pawn_Credit)
        End With


        Return tempDs

    End Function


    Public Function UpdatePawnCredit(ByVal PawnCreditSet As DataSet)

        If PawnCreditSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If PawnCreditSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Pawn_Credit)

        With dsCommand_Pawn_Credit
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(PawnCreditSet, Table_Pawn_Credit)


        End With

        PawnCreditSet.AcceptChanges()

    End Function

    
    

End Class

