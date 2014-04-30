Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class OppositeGuaranteeAssurer
    Public Const Table_OppositeGuaranteeAssurer As String = "opposite_guarantee_assurer"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_OppositeGuaranteeAssurer As SqlDataAdapter

    '定义查询命令
    Private GetOppositeGuaranteeAssurerInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_OppositeGuaranteeAssurer = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetOppositeGuaranteeAssurerInfo("null")
    End Sub

    '获取意向书信息
    Public Function GetOppositeGuaranteeAssurerInfo(ByVal strSQL_Condition_OppositeGuaranteeAssurer As String) As DataSet

        Dim tempDs As New DataSet()

        If GetOppositeGuaranteeAssurerInfoCommand Is Nothing Then

            GetOppositeGuaranteeAssurerInfoCommand = New SqlCommand("GetOppositeGuaranteeAssurerInfo", conn)
            GetOppositeGuaranteeAssurerInfoCommand.CommandType = CommandType.StoredProcedure
            GetOppositeGuaranteeAssurerInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_OppositeGuaranteeAssurer
            .SelectCommand = GetOppositeGuaranteeAssurerInfoCommand
            .SelectCommand.Transaction = ts
            GetOppositeGuaranteeAssurerInfoCommand.Parameters("@Condition").Value = strSQL_Condition_OppositeGuaranteeAssurer
            .Fill(tempDs, Table_OppositeGuaranteeAssurer)
        End With

        Return tempDs

    End Function

    '更新意向书信息
    Public Function UpdateOppositeGuaranteeAssurer(ByVal OppositeGuaranteeAssurerSet As DataSet)

        If OppositeGuaranteeAssurerSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If OppositeGuaranteeAssurerSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_OppositeGuaranteeAssurer)

        With dsCommand_OppositeGuaranteeAssurer
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(OppositeGuaranteeAssurerSet, Table_OppositeGuaranteeAssurer)

        End With

        OppositeGuaranteeAssurerSet.AcceptChanges()

    End Function
End Class
