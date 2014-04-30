
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplEndProject
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义项目对象引用
    Private project As project

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTask As WfProjectTask
    Private WfProjectMessages As WfProjectMessages
    Private WfProjectTimingTask As WfProjectTimingTask
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTrack As WfProjectTrack

    Private WorkLog As Worklog

    Private TimingServer As TimingServer

    Private CommonQuery As CommonQuery

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化项目对象
        project = New Project(conn, ts)

        '实例化参与人对象引用
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectMessages = New WfProjectMessages(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTrack = New WfProjectTrack(conn, ts)

        WorkLog = New WorkLog(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)

        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim i As Integer

        '设置项目结束标识isliving=0
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsTemp, dsTempProject, dsTempAttend, dsTempTrans, dsMsg, dsTempTask, dsTempTimingTask, dsProjectTrack As DataSet
        Dim tmpStatus, tmpAttend As String
        dsTempProject = project.GetProjectInfo(strSql)

        '异常处理  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        dsTempProject.Tables(0).Rows(0).Item("isliving") = 0
        project.UpdateProject(dsTempProject)

        '删除该项目的所有消息
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsMsg = WfProjectMessages.GetWfProjectMessagesInfo(strSql)

        For i = 0 To dsMsg.Tables(0).Rows.Count - 1
            dsMsg.Tables(0).Rows(i).Delete()
        Next
        WfProjectMessages.UpdateWfProjectMessages(dsMsg)


        'qxd modify 2004-10-25 据深圳秦工提出：项目结束的信息只发给项目经理！-------------------start
        ''给参与过该项目的所有人员发送项目结束的消息
        'strSql = " select distinct attend_person from project_task_attendee" & _
        '         " where project_code=" & "'" & projectID & "'" & _
        '         " and attend_person<>''"
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    tmpAttend = dsTemp.Tables(0).Rows(i).Item("attend_person")
        '    TimingServer.AddMsg(workFlowID, projectID, taskID, tmpAttend, "8", "N")
        'Next

        '给参与过该项目的所有人员发送项目结束的消息 -> 项目结束的信息只发给项目经理
        'strSql = " select distinct attend_person from project_task_attendee" & _
        '         " where project_code=" & "'" & projectID & "'" & _
        '         " and attend_person<>'' and role_id in ('24') "
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    tmpAttend = dsTemp.Tables(0).Rows(i).Item("attend_person")
        '    TimingServer.AddMsg(workFlowID, projectID, taskID, tmpAttend, "8", "N")
        'Next
        'qxd modify 2004-10-25 据深圳秦工提出：项目结束的信息只发给项目经理！-------------------end

        ''将项目所有正在处理的任务置为""
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='P'}"
        'Dim dsTaskAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTaskAttend.Tables(0).Rows.Count - 1
        '    dsTaskAttend.Tables(0).Rows(i).Item("task_status") = ""
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTaskAttend)

        ''在定时任务表将与参数匹配的提示任务状态置为“E”；
        'strSql = "{project_code=" & "'" & projectID & "'" & "}"
        'dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''   所有地提示任务状态置为“E”
        'For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
        '    dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        'Next
        'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '删除任务参与人表中项目编码的任务；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '删除转移表中的明细记录
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        For i = 0 To dsTempTrans.Tables(0).Rows.Count - 1
            dsTempTrans.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTrans)

        '删除定时任务中指定项目编码的定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '删除任务跟踪表中的记录
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)
        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1
            dsProjectTrack.Tables(0).Rows(i).Delete()
        Next
        WfProjectTrack.UpdateWfProjectTrack(dsTempTimingTask)

        '删除任务表中的任务；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
        For i = 0 To dsTempTask.Tables(0).Rows.Count - 1

            dsTempTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTask.UpdateWfProjectTask(dsTempTask)

    End Function

End Class
