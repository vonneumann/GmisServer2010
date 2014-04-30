Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplStartAferGuaranteeTask
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义项目帐户明细对象引用
    Private ProjectAccountDetail As ProjectAccountDetail

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private TracePlan As TracePlan


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化项目帐户明细对象
        ProjectAccountDetail = New ProjectAccountDetail(conn, ts)

        '实例化定时任务对象
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        TracePlan = New TracePlan(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim i As Integer
        Dim strSql As String
        Dim dsTempAttend As DataSet

        '将还款登记任务置'P'(因为还款登记的定时任务的已经建立并开始,不需要再走StartupTask过程)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundRecord'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '将登记保后跟踪记录任务置'P'(因为登记保后跟踪记录的定时任务的已经建立并开始,不需要再走StartupTask过程)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordProjectTraceInfo'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function
End Class
