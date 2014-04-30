
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplEndProject
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '������Ŀ��������
    Private project As project

    '��������˶�������
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


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������Ŀ����
        project = New Project(conn, ts)

        'ʵ���������˶�������
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

        '������Ŀ������ʶisliving=0
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsTemp, dsTempProject, dsTempAttend, dsTempTrans, dsMsg, dsTempTask, dsTempTimingTask, dsProjectTrack As DataSet
        Dim tmpStatus, tmpAttend As String
        dsTempProject = project.GetProjectInfo(strSql)

        '�쳣����  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        dsTempProject.Tables(0).Rows(0).Item("isliving") = 0
        project.UpdateProject(dsTempProject)

        'ɾ������Ŀ��������Ϣ
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsMsg = WfProjectMessages.GetWfProjectMessagesInfo(strSql)

        For i = 0 To dsMsg.Tables(0).Rows.Count - 1
            dsMsg.Tables(0).Rows(i).Delete()
        Next
        WfProjectMessages.UpdateWfProjectMessages(dsMsg)


        'qxd modify 2004-10-25 �������ع��������Ŀ��������Ϣֻ������Ŀ����-------------------start
        ''�����������Ŀ��������Ա������Ŀ��������Ϣ
        'strSql = " select distinct attend_person from project_task_attendee" & _
        '         " where project_code=" & "'" & projectID & "'" & _
        '         " and attend_person<>''"
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    tmpAttend = dsTemp.Tables(0).Rows(i).Item("attend_person")
        '    TimingServer.AddMsg(workFlowID, projectID, taskID, tmpAttend, "8", "N")
        'Next

        '�����������Ŀ��������Ա������Ŀ��������Ϣ -> ��Ŀ��������Ϣֻ������Ŀ����
        'strSql = " select distinct attend_person from project_task_attendee" & _
        '         " where project_code=" & "'" & projectID & "'" & _
        '         " and attend_person<>'' and role_id in ('24') "
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    tmpAttend = dsTemp.Tables(0).Rows(i).Item("attend_person")
        '    TimingServer.AddMsg(workFlowID, projectID, taskID, tmpAttend, "8", "N")
        'Next
        'qxd modify 2004-10-25 �������ع��������Ŀ��������Ϣֻ������Ŀ����-------------------end

        ''����Ŀ�������ڴ����������Ϊ""
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='P'}"
        'Dim dsTaskAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTaskAttend.Tables(0).Rows.Count - 1
        '    dsTaskAttend.Tables(0).Rows(i).Item("task_status") = ""
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTaskAttend)

        ''�ڶ�ʱ����������ƥ�����ʾ����״̬��Ϊ��E����
        'strSql = "{project_code=" & "'" & projectID & "'" & "}"
        'dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''   ���е���ʾ����״̬��Ϊ��E��
        'For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
        '    dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        'Next
        'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        'ɾ����������˱�����Ŀ���������
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        'ɾ��ת�Ʊ��е���ϸ��¼
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        For i = 0 To dsTempTrans.Tables(0).Rows.Count - 1
            dsTempTrans.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTrans)

        'ɾ����ʱ������ָ����Ŀ����Ķ�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        'ɾ��������ٱ��еļ�¼
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)
        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1
            dsProjectTrack.Tables(0).Rows(i).Delete()
        Next
        WfProjectTrack.UpdateWfProjectTrack(dsTempTimingTask)

        'ɾ��������е�����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
        For i = 0 To dsTempTask.Tables(0).Rows.Count - 1

            dsTempTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTask.UpdateWfProjectTask(dsTempTask)

    End Function

End Class
