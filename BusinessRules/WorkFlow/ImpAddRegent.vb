Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'添加评委任务
Public Class ImpAddRegent
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义评审会对象引用
    Private ConfTrial As ConfTrial

    '定义会议对象引用
    Private Conference As Conference

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

    '定义消息对象引用
    Private WfProjectMessages As WfProjectMessages




    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans


        '实例化评审会对象
        ConfTrial = New ConfTrial(conn, ts)

        '实例化会议对象
        Conference = New Conference(conn, ts)

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        '实例化定时任务对象
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        '实例化消息对象
        WfProjectMessages = New WfProjectMessages(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i, j, iCommitteemanCount, iAttendCount As Integer
        Dim dsTempCommitOpinion, dsTempConfCode, dsTempConference, dsTempAttend, dsTempTaskMessages, dsTempTimingTask, dsTempDeleteAttend, dsTempRoleTemplate As DataSet
        Dim tmpTrialTime, tmpConfCode As Integer
        Dim tmpConfDate As DateTime
        Dim tmpCommitteeman, tmpWorkflowID As String
        Dim newRow As DataRow
        '①	在Committeeman-Opinion表获取与ProjectId匹配且OPINION为空的所有评委和trial-times；
        strSql = "{project_code=" & "'" & projectID & "'" & " and opinion is null" & "}"
        dsTempCommitOpinion = ConfTrial.GetConfTrialInfo("null", strSql)

        '异常处理  
        If dsTempCommitOpinion.Tables(1).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempCommitOpinion.Tables(1))
            Throw wfErr
        End If

        tmpTrialTime = dsTempCommitOpinion.Tables(1).Rows(0).Item("trial_times")
        iCommitteemanCount = dsTempCommitOpinion.Tables(1).Rows.Count

        '②	在conference-trail表获取与ProjectId、trial-times匹配的conference-code;
        strSql = "{project_code=" & "'" & projectID & "'" & " and trial_times=" & tmpTrialTime & "}"
        dsTempConfCode = ConfTrial.GetConfTrialInfo(strSql, "null")

        '异常处理  
        If dsTempConfCode.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempConfCode.Tables(0))
            Throw wfErr
        End If

        tmpConfCode = dsTempConfCode.Tables(0).Rows(0).Item("conference_code")

        '③	在conference表获取与conference-code匹配的conference-date;
        strSql = "{conference_code=" & tmpConfCode & "}"
        dsTempConference = Conference.GetConferenceInfo(strSql, "null")

        '异常处理  
        If dsTempConference.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempConference.Tables(0))
            Throw wfErr
        End If

        tmpConfDate = dsTempConference.Tables(0).Rows(0).Item("conference_date")


        '先获取当前任务的模版ID(作为复制模版的实例时的限制)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
        If dsTempDeleteAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempDeleteAttend.Tables(0))
            Throw wfErr
        End If

        tmpWorkflowID = dsTempDeleteAttend.Tables(0).Rows(0).Item("workflow_id")

        '先删除任务角色表中原有评委的任务（上一次分配的评委）
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='26'" & "}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempDeleteAttend.Tables(0).Rows.Count - 1
            dsTempDeleteAttend.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        '再复制任务角色模版的评委的任务
        strSql = "{workflow_id=" & "'" & tmpWorkflowID & "'" & " and  role_id='26'" & "}"
        Dim WfTaskRoleTemplate As New WfTaskRoleTemplate(conn, ts)
        dsTempRoleTemplate = WfTaskRoleTemplate.GetWfTaskRoleTemplateInfo(strSql)

        For i = 0 To dsTempRoleTemplate.Tables(0).Rows.Count - 1
            newRow = dsTempDeleteAttend.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = tmpWorkflowID
                .Item("project_code") = projectID    '将所有添加任务的工作流ID置为项目编码
                .Item("task_id") = dsTempRoleTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTempRoleTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTempDeleteAttend.Tables(0).Rows.Add(newRow)
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='26'" & "}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        iAttendCount = dsTempAttend.Tables(0).Rows.Count

        '④	对于每位评委
        ' 在项目的任务参与人表复制role-id=26的任务（n-1） ,将参与人改为评委；
        If iCommitteemanCount <> 0 Then

            For i = 0 To iAttendCount - 1
                dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsTempCommitOpinion.Tables(1).Rows(0).Item("committeeman")
                '判断是否是多个任务
                If iAttendCount > 1 Then
                    For j = 1 To iCommitteemanCount - 1
                        newRow = dsTempAttend.Tables(0).NewRow
                        With dsTempAttend.Tables(0).Rows(i)
                            newRow.Item("project_code") = .Item("project_code")
                            newRow.Item("workflow_id") = .Item("workflow_id")
                            newRow.Item("task_id") = .Item("task_id")
                            newRow.Item("role_id") = .Item("role_id")
                            newRow.Item("attend_person") = dsTempCommitOpinion.Tables(1).Rows(j).Item("committeeman")
                        End With
                        dsTempAttend.Tables(0).Rows.Add(newRow)
                    Next
                End If
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
            '向消息库插入消息（消息名=开会日期（conference-date）出席评审会，参与人=评委,确认标志=“N”）；
            For i = 0 To dsTempCommitOpinion.Tables(1).Rows.Count - 1
                tmpCommitteeman = dsTempCommitOpinion.Tables(1).Rows(i).Item("committeeman")
                newRow = dsTempTaskMessages.Tables(0).NewRow
                With newRow
                    .Item("project_code") = projectID
                    .Item("message_content") = CStr(tmpConfDate) & " 出席评审会"
                    .Item("accepter") = tmpCommitteeman
                    .Item("send_time") = Now
                    .Item("is_affirmed") = "N"
                End With
                dsTempTaskMessages.Tables(0).Rows.Add(newRow)
            Next
            WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
        Else

            '抛出未分配评委错误
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowAddRegentErr()
            Throw wfErr

        End If

        '⑤	在定时任务表匹配workflow_id,project_id,task_id =ReviewMeeting的任务，将定时任务的开始时间改为conference-date；
        '⑥	将定时任务状态置为“P”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ReviewMeeting'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Item("start_time") = tmpConfDate
            dsTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

    End Function
End Class
