Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ReturnReceipt

    Public Const Table_ReturnReceipt As String = "loan_return_receipt"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_ReturnReceipt As SqlDataAdapter

    '定义查询命令
    Private GetReturnReceiptInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_ReturnReceipt = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetReturnReceiptInfo("null")
    End Sub

    '获取放款回执信息
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

    '更新放款回执信息
    Public Function UpdateReturnReceipt(ByVal ReturnReceiptSet As DataSet)

        If ReturnReceiptSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
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
