Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class SignaturePlan
    Public Const Table_Signature_Plan As String = "signature_plan"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_SignaturePlan As SqlDataAdapter

    '定义查询命令
    Private GetSignaturePlanInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_SignaturePlan = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetSignaturePlanInfo("null")
    End Sub

    '获取签约计划信息
    Public Function GetSignaturePlanInfo(ByVal strSQL_Condition_SignaturePlan As String) As DataSet

        Dim tempDs As New DataSet()

        If GetSignaturePlanInfoCommand Is Nothing Then

            GetSignaturePlanInfoCommand = New SqlCommand("GetSignaturePlanInfo", conn)
            GetSignaturePlanInfoCommand.CommandType = CommandType.StoredProcedure
            GetSignaturePlanInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_SignaturePlan
            .SelectCommand = GetSignaturePlanInfoCommand
            .SelectCommand.Transaction = ts
            GetSignaturePlanInfoCommand.Parameters("@Condition").Value = strSQL_Condition_SignaturePlan
            .Fill(tempDs, Table_Signature_Plan)
        End With

        Return tempDs

    End Function


    '更新签约计划信息
    Public Function UpdateSignaturePlan(ByVal SignaturePlanSet As DataSet)

        If SignaturePlanSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If SignaturePlanSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_SignaturePlan)

        With dsCommand_SignaturePlan
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(SignaturePlanSet, Table_Signature_Plan)

            SignaturePlanSet.AcceptChanges()
        End With


    End Function
End Class
