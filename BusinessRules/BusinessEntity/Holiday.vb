Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Holiday
    Public Const Table_Holiday As String = "holiday"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Holiday As SqlDataAdapter

    '定义查询命令
    Private GetHolidayInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Holiday = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetHolidayInfo("null")
    End Sub

    '获取假期信息
    Public Function GetHolidayInfo(ByVal strSQL_Condition_Holiday As String) As DataSet

        Dim tempDs As New DataSet()

        If GetHolidayInfoCommand Is Nothing Then

            GetHolidayInfoCommand = New SqlCommand("GetHolidayInfo", conn)
            GetHolidayInfoCommand.CommandType = CommandType.StoredProcedure
            GetHolidayInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Holiday
            .SelectCommand = GetHolidayInfoCommand
            .SelectCommand.Transaction = ts
            GetHolidayInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Holiday
            .Fill(tempDs, Table_Holiday)
        End With

        Return tempDs

    End Function

    '更新假期信息
    Public Function UpdateHoliday(ByVal HolidaySet As DataSet)

        If HolidaySet Is Nothing Then
            Exit Function
        End If


        '如果记录集未发生任何变化，则退出过程
        If HolidaySet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Holiday)

        With dsCommand_Holiday
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(HolidaySet, Table_Holiday)

        End With

        HolidaySet.AcceptChanges()
    End Function

End Class
