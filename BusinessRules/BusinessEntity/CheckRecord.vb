Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CheckRecord
    Public Const Table_CheckRecord As String = "guarantee_check_record"
    Public Const Table_CheckAlarm As String = "guarantee_check_alarm"


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_CheckRecord As SqlDataAdapter
    Private dsCommand_CheckAlarm As SqlDataAdapter

    '定义查询命令
    Private GetCheckRecordInfoCommand As SqlCommand
    Private GetMaxCheckRecordNumCommand As SqlCommand
    Private GetCheckAlarmInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_CheckRecord = New SqlDataAdapter()
        dsCommand_CheckAlarm = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetCheckRecordInfo("null", "null")

    End Sub

    '获取保后检查记录信息
    Public Function GetCheckRecordInfo(ByVal strSQL_Condition_CheckRecord As String, ByVal strSQL_Condition_CheckAlarm As String) As DataSet

        Dim tempDs As New DataSet()

        If GetCheckRecordInfoCommand Is Nothing Then

            GetCheckRecordInfoCommand = New SqlCommand("GetCheckRecordInfo", conn)
            GetCheckRecordInfoCommand.CommandType = CommandType.StoredProcedure
            GetCheckRecordInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CheckRecord
            .SelectCommand = GetCheckRecordInfoCommand
            .SelectCommand.Transaction = ts
            GetCheckRecordInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CheckRecord
            .Fill(tempDs, Table_CheckRecord)
        End With

        If GetCheckAlarmInfoCommand Is Nothing Then

            GetCheckAlarmInfoCommand = New SqlCommand("GetCheckAlarmInfo", conn)
            GetCheckAlarmInfoCommand.CommandType = CommandType.StoredProcedure
            GetCheckAlarmInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CheckAlarm
            .SelectCommand = GetCheckAlarmInfoCommand
            .SelectCommand.Transaction = ts
            GetCheckAlarmInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CheckAlarm
            .Fill(tempDs, Table_CheckAlarm)
        End With

        GetCheckRecordInfo = tempDs

    End Function

    '获取最大序列号
    Public Function GetMaxCheckRecordNum(ByVal projectID As String) As Integer

        If GetMaxCheckRecordNumCommand Is Nothing Then

            GetMaxCheckRecordNumCommand = New SqlCommand("GetMaxCheckRecordNum", conn)
            GetMaxCheckRecordNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxCheckRecordNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxCheckRecordNumCommand.Parameters.Add(New SqlParameter("@maxCheckRecordNum", SqlDbType.Int))
            GetMaxCheckRecordNumCommand.Parameters.Item("@maxCheckRecordNum").Direction = ParameterDirection.Output
            GetMaxCheckRecordNumCommand.Transaction = ts
        End If

        GetMaxCheckRecordNumCommand.Parameters("@projectID").Value = projectID
        GetMaxCheckRecordNumCommand.ExecuteNonQuery()
        GetMaxCheckRecordNum = GetMaxCheckRecordNumCommand.Parameters.Item("@maxCheckRecordNum").Value
    End Function

    '更新保后检查记录信息
    Private Function UpdateCheckRecord(ByVal CheckRecordSet As DataSet)
        Try
            Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CheckRecord)


            With dsCommand_CheckRecord
                .InsertCommand = bd.GetInsertCommand
                .UpdateCommand = bd.GetUpdateCommand
                .DeleteCommand = bd.GetDeleteCommand


                .InsertCommand.Transaction = ts
                .UpdateCommand.Transaction = ts
                .DeleteCommand.Transaction = ts
                .Update(CheckRecordSet, Table_CheckRecord)

                CheckRecordSet.AcceptChanges()
        End With
        Catch
            MsgBox(Err.Description)
        End Try

    End Function

    '更新保后预警记录信息
    Private Function UpdateCheckAlarm(ByVal CheckAlarmSet As DataSet)


        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CheckAlarm)

        With dsCommand_CheckAlarm
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts
            .Update(CheckAlarmSet, Table_CheckAlarm)

            CheckAlarmSet.AcceptChanges()
        End With

    End Function

    Public Function UpdateCheckRecordAlarm(ByVal CheckRecordAlarmSet As DataSet)

        If CheckRecordAlarmSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If CheckRecordAlarmSet.HasChanges = False Then
            Exit Function
        End If


        '删除操作
        If IsNothing(CheckRecordAlarmSet.GetChanges(DataRowState.Deleted)) = False Then
            '先删明细表，再删主表
            UpdateCheckAlarm(CheckRecordAlarmSet.GetChanges(DataRowState.Deleted))
            UpdateCheckRecord(CheckRecordAlarmSet.GetChanges(DataRowState.Deleted))

        End If

        '新增作
        If IsNothing(CheckRecordAlarmSet.GetChanges(DataRowState.Added)) = False Then
            UpdateCheckRecord(CheckRecordAlarmSet.GetChanges(DataRowState.Added))
            UpdateCheckAlarm(CheckRecordAlarmSet.GetChanges(DataRowState.Added))
        End If

        '更新操作
        If IsNothing(CheckRecordAlarmSet.GetChanges(DataRowState.Modified)) = False Then
            UpdateCheckRecord(CheckRecordAlarmSet.GetChanges(DataRowState.Modified))
            UpdateCheckAlarm(CheckRecordAlarmSet.GetChanges(DataRowState.Modified))
        End If

        CheckRecordAlarmSet.AcceptChanges()

    End Function

End Class
