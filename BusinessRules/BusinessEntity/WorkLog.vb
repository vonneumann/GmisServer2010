Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WorkLog

    Public Const Table_Work_Log As String = "work_log"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WorkLog As SqlDataAdapter

    '定义查询命令
    Private GetWorkLogInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction


    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_WorkLog = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetWorkLogInfo("null")
    End Sub

    '获取工作记录信息
    Public Function GetWorkLogInfo(ByVal strSQL_Condition_WorkLog As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWorkLogInfoCommand Is Nothing Then

            GetWorkLogInfoCommand = New SqlCommand("GetWorkLogInfo", conn)
            GetWorkLogInfoCommand.CommandType = CommandType.StoredProcedure
            GetWorkLogInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WorkLog
            .SelectCommand = GetWorkLogInfoCommand
            .SelectCommand.Transaction = ts
            GetWorkLogInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WorkLog
            .Fill(tempDs, Table_Work_Log)
        End With

        Return tempDs


    End Function

    '更新工作记录信息
    Public Function UpdateWorkLog(ByVal WorkLogSet As DataSet)

        If WorkLogSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If WorkLogSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WorkLog)

        With dsCommand_WorkLog
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            '    '获取记录次数
            '    Dim tmpProjectID As String = WorkLogSet.Tables(0).Rows(0).Item("project_code")
            '    Dim strSql As String = "{project_code=" & "'" & tmpProjectID & "'" & " order by serial_num}"
            '    Dim dsTemp As DataSet = GetWorkLogInfo(strSql)
            '    Dim tmpSerialNum As Integer
            '    Dim tmpRowNum As Integer = dsTemp.Tables(0).Rows.Count
            '    If tmpRowNum = 0 Then
            '        '如果找不到记录，则编号从1开始
            '        tmpSerialNum = 1
            '    Else
            '        '   获取当前项目的最大次数＋1
            '        tmpSerialNum = dsTemp.Tables(0).Rows(tmpRowNum - 1).Item("serial_num") + 1
            '    End If

            '    Dim i As Integer
            '    For i = 0 To WorkLogSet.Tables(0).Rows.Count - 1
            '        WorkLogSet.Tables(0).Rows(i).Item("serial_num") = tmpSerialNum
            '        tmpSerialNum = tmpSerialNum + 1
            '    Next

            .Update(WorkLogSet, Table_Work_Log)

            WorkLogSet.AcceptChanges()
        End With


    End Function
End Class
