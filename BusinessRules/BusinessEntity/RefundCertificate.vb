Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class RefundCertificate

    Public Const Table_RefundCertificate As String = "refund_certificate"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_RefundCertificate As SqlDataAdapter

    '定义查询命令
    Private GetRefundCertificateInfoCommand As SqlCommand
    Private GetMaxRefundCertificateNumCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_RefundCertificate = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetRefundCertificateInfo("null")
    End Sub

    '获取还款证明书信息
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

    '更新还款证明书信息
    Public Function UpdateRefundCertificate(ByVal RefundCertificateSet As DataSet)

        If RefundCertificateSet Is Nothing Then
            Exit Function
        End If



        '如果记录集未发生任何变化，则退出过程
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
