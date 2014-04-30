Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

''启动保后登记任务
'关闭登记保后检查记录表、审核保后检查记录表、项目评价任务；
'停止登记保后检查记录表、登记还款定时任务
Public Class ImplStartSubsequentReg
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义定时任务
    Private WfProjectTimingTask As WfProjectTimingTask

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTempTaskAttendee As DataSet
        Dim i As Integer

        ''①	将登记逾期追踪信息任务（TID=OverdueTrailRecord）所有角色的任务状态置为“P”；
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='OverdueTrailRecord'}"
        'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
        '    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        ''②	将登记代偿信息任务（TID=RefundDebtInfo）所有角色的任务状态置为“P”；
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundDebtInfo'}"
        'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
        '    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        'Next

        '1.关闭登记保后检查记录表、审核保后检查记录表、项目评价任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and (task_id='RecordProjectTraceInfo' or task_id='CheckProjectProcess' or task_id='AppraiseProjectProcess')}"

        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        '2.停止登记保后检查记录表、登记还款定时任务
        Dim dsTemp As DataSet
        strSql = "{project_code='" & projectID & "' and (task_id='RecordProjectTraceInfo' or task_id='RefundRecord')}"
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

    End Function

End Class
