
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class DdGuarantyStatus

    Public Const Table_DdGuarantyStatus As String = "dd_guaranty_status"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_DdGuarantyStatus As SqlDataAdapter

    '定义查询命令
    Private GetDdGuarantyStatusInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_DdGuarantyStatus = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetDdGuarantyStatusInfo("null")
    End Sub

    '获取项目进度信息
    Public Function GetDdGuarantyStatusInfo(ByVal strSQL_Condition_DdGuarantyStatus As String) As DataSet

        Dim tempDs As New DataSet()

        If GetDdGuarantyStatusInfoCommand Is Nothing Then

            GetDdGuarantyStatusInfoCommand = New SqlCommand("GetDdGuarantyStatusInfo", conn)
            GetDdGuarantyStatusInfoCommand.CommandType = CommandType.StoredProcedure
            GetDdGuarantyStatusInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_DdGuarantyStatus
            .SelectCommand = GetDdGuarantyStatusInfoCommand
            .SelectCommand.Transaction = ts
            GetDdGuarantyStatusInfoCommand.Parameters("@Condition").Value = strSQL_Condition_DdGuarantyStatus
            .Fill(tempDs, Table_DdGuarantyStatus)
        End With

        Return tempDs

    End Function

    '更新项目进度信息
    Public Function UpdateDdGuarantyStatus(ByVal DdGuarantyStatusSet As DataSet)

        If DdGuarantyStatusSet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If DdGuarantyStatusSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_DdGuarantyStatus)

        With dsCommand_DdGuarantyStatus
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(DdGuarantyStatusSet, Table_DdGuarantyStatus)

        End With

        DdGuarantyStatusSet.AcceptChanges()
    End Function
End Class
