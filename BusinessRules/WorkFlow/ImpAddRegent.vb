Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�����ί����
Public Class ImpAddRegent
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����������������
    Private ConfTrial As ConfTrial

    '��������������
    Private Conference As Conference

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���嶨ʱ�����������
    Private WfProjectTimingTask As WfProjectTimingTask

    '������Ϣ��������
    Private WfProjectMessages As WfProjectMessages




    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans


        'ʵ������������
        ConfTrial = New ConfTrial(conn, ts)

        'ʵ�����������
        Conference = New Conference(conn, ts)

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        'ʵ������ʱ�������
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        'ʵ������Ϣ����
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
        '��	��Committeeman-Opinion���ȡ��ProjectIdƥ����OPINIONΪ�յ�������ί��trial-times��
        strSql = "{project_code=" & "'" & projectID & "'" & " and opinion is null" & "}"
        dsTempCommitOpinion = ConfTrial.GetConfTrialInfo("null", strSql)

        '�쳣����  
        If dsTempCommitOpinion.Tables(1).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempCommitOpinion.Tables(1))
            Throw wfErr
        End If

        tmpTrialTime = dsTempCommitOpinion.Tables(1).Rows(0).Item("trial_times")
        iCommitteemanCount = dsTempCommitOpinion.Tables(1).Rows.Count

        '��	��conference-trail���ȡ��ProjectId��trial-timesƥ���conference-code;
        strSql = "{project_code=" & "'" & projectID & "'" & " and trial_times=" & tmpTrialTime & "}"
        dsTempConfCode = ConfTrial.GetConfTrialInfo(strSql, "null")

        '�쳣����  
        If dsTempConfCode.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempConfCode.Tables(0))
            Throw wfErr
        End If

        tmpConfCode = dsTempConfCode.Tables(0).Rows(0).Item("conference_code")

        '��	��conference���ȡ��conference-codeƥ���conference-date;
        strSql = "{conference_code=" & tmpConfCode & "}"
        dsTempConference = Conference.GetConferenceInfo(strSql, "null")

        '�쳣����  
        If dsTempConference.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempConference.Tables(0))
            Throw wfErr
        End If

        tmpConfDate = dsTempConference.Tables(0).Rows(0).Item("conference_date")


        '�Ȼ�ȡ��ǰ�����ģ��ID(��Ϊ����ģ���ʵ��ʱ������)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsTempDeleteAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempDeleteAttend.Tables(0))
            Throw wfErr
        End If

        tmpWorkflowID = dsTempDeleteAttend.Tables(0).Rows(0).Item("workflow_id")

        '��ɾ�������ɫ����ԭ����ί��������һ�η������ί��
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='26'" & "}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempDeleteAttend.Tables(0).Rows.Count - 1
            dsTempDeleteAttend.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        '�ٸ��������ɫģ�����ί������
        strSql = "{workflow_id=" & "'" & tmpWorkflowID & "'" & " and  role_id='26'" & "}"
        Dim WfTaskRoleTemplate As New WfTaskRoleTemplate(conn, ts)
        dsTempRoleTemplate = WfTaskRoleTemplate.GetWfTaskRoleTemplateInfo(strSql)

        For i = 0 To dsTempRoleTemplate.Tables(0).Rows.Count - 1
            newRow = dsTempDeleteAttend.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = tmpWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
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

        '��	����ÿλ��ί
        ' ����Ŀ����������˱���role-id=26������n-1�� ,�������˸�Ϊ��ί��
        If iCommitteemanCount <> 0 Then

            For i = 0 To iAttendCount - 1
                dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsTempCommitOpinion.Tables(1).Rows(0).Item("committeeman")
                '�ж��Ƿ��Ƕ������
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
            '����Ϣ�������Ϣ����Ϣ��=�������ڣ�conference-date����ϯ����ᣬ������=��ί,ȷ�ϱ�־=��N������
            For i = 0 To dsTempCommitOpinion.Tables(1).Rows.Count - 1
                tmpCommitteeman = dsTempCommitOpinion.Tables(1).Rows(i).Item("committeeman")
                newRow = dsTempTaskMessages.Tables(0).NewRow
                With newRow
                    .Item("project_code") = projectID
                    .Item("message_content") = CStr(tmpConfDate) & " ��ϯ�����"
                    .Item("accepter") = tmpCommitteeman
                    .Item("send_time") = Now
                    .Item("is_affirmed") = "N"
                End With
                dsTempTaskMessages.Tables(0).Rows.Add(newRow)
            Next
            WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
        Else

            '�׳�δ������ί����
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowAddRegentErr()
            Throw wfErr

        End If

        '��	�ڶ�ʱ�����ƥ��workflow_id,project_id,task_id =ReviewMeeting�����񣬽���ʱ����Ŀ�ʼʱ���Ϊconference-date��
        '��	����ʱ����״̬��Ϊ��P����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ReviewMeeting'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Item("start_time") = tmpConfDate
            dsTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

    End Function
End Class
