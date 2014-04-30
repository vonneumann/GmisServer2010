Option Explicit On 

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'��������������ÿ��������ÿ��7��00ʱ�����÷���TimingServer�����������ʱ����
Public Class TimingServer

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���嶨ʱ�����������
    Private WfProjectTimingTask As WfProjectTimingTask

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���������б�����
    Private WfProjectTask As WfProjectTask

    ''���幤����¼��������
    'Private WorkLog As WorkLog

    '������Ŀ��������
    Private Project As Project

    '������Ϣ�ֵ��������
    Private WfMessagesTemplate As WfMessagesTemplate

    '���幤������Ϣ��������
    Private WfProjectMessages As WfProjectMessages

    '�����ɫԱ����������
    Private role As role

    '���幤������������
    Private WorkFlow As WorkFlow

    ''����һ���µĶ�ʱ��������,�����ͱ���
    'Private Ddone(30) As Boolean

    '������ڶ�������
    Private Holiday As Holiday

    Private TimeServerLog As TimeServerLog

    Private Branch As Branch

    Private CommonQuery As CommonQuery

    '���嵱�죬��Сʱɨ��Ĳ�������
    Private bDay As Boolean
    Private bHour As Boolean

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction, ByVal isDayScan As Boolean, ByVal isHourScan As Boolean)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������ʱ�������
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        'ʵ���������б����
        WfProjectTask = New WfProjectTask(conn, ts)

        'ʵ������Ŀ����
        Project = New Project(conn, ts)

        ''ʵ����������¼����
        'WorkLog = New WorkLog(conn, ts)

        'ʵ������Ϣ�ֵ����
        WfMessagesTemplate = New WfMessagesTemplate(conn, ts)

        'ʵ������������Ϣ����
        WfProjectMessages = New WfProjectMessages(conn, ts)

        'ʵ������ɫԱ������
        role = New Role(conn, ts)

        'ʵ������������Ϣ����
        WorkFlow = New WorkFlow(conn, ts)

        'ʵ�������ڶ���
        Holiday = New Holiday(conn, ts)

        TimeServerLog = New TimeServerLog(conn, ts)


        Branch = New Branch(conn, ts)

        CommonQuery = New CommonQuery(conn, ts)

        '������ֵ���赱�죬��Сʱɨ��Ĳ�������
        bDay = isDayScan
        bHour = isHourScan

    End Sub

    Public Function TimingServer()
        Dim i, j As Integer
        Dim strSql As String
        Dim tmpStartTime As DateTime
        Dim tmpPromptCount, tmpRecordCount, NoScanedDay As Integer
        Dim dsTempTimingTask, dsTempTaskAttenddee, dsTempTask, dsTempWorkLog, dsTemp, dsHoliday, dsTempProject As DataSet

        Dim dsTimeServerLog As DataSet = TimeServerLog.GetTimeServerLogInfo("null")
        Dim drTimeServerLog As DataRow

        Dim tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpMessageID, tmpStatus, tmpRoleID, tmpStaffID, tmpTaskPhase, tmpTaskStatus, tmpHeader As String

        Dim tmpDistance As Integer

        '������־�ļ�
        'Dim fs As FileStream = New FileStream("Log.txt", FileMode.OpenOrCreate) '�򿪲�������־�ļ�
        'Dim w As StreamWriter = New StreamWriter(fs) '����д�ļ���
        'w.BaseStream.Seek(0, SeekOrigin.End) 'ָ���ļ�ĩβ

        '��	��ȡϵͳ���ڣ�
        Dim sysTime As DateTime = Now

        '�жϵ����Ƿ���Ҫɨ�裨���������Ϣ������

        If bDay = False Then


            '��ȡ���ڵ���Today��NoWorkingDay. Scaned=False������NoScanedDay��
            Dim sysDay As String = FormatDateTime(sysTime, DateFormat.ShortDate)
            strSql = "{holiday<=" & "'" & sysDay & "'" & " and isnull(scaned,0)<>1}"
            dsHoliday = Holiday.GetHolidayInfo(strSql)
            NoScanedDay = dsHoliday.Tables(0).Rows.Count

            '������Today��δɨ���NoWorkingDay��Scaned��ΪTrue;
            For i = 0 To dsHoliday.Tables(0).Rows.Count - 1

                dsHoliday.Tables(0).Rows(i).Item("scaned") = 1

            Next
            Holiday.UpdateHoliday(dsHoliday)

            '����ʱ���������������Ϊ��A����״̬Ϊ��P���Ķ�ʱ����ʼʱ���Ϊ��ʼʱ��+NoScanedDay��24;
            If NoScanedDay <> 0 Then
                strSql = "{type='A' and status='P'}"
                dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                    dsTemp.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Hour, NoScanedDay * 24, dsTemp.Tables(0).Rows(i).Item("start_time"))
                Next
                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)
            End If


            '��	��ȡ��ʱ����״̬Ϊ��P��,��������Ϊ��T,P,W��������
            strSql = "{status='P' and type in ('T','P','W','R','M')}"
            dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

            '	����ÿһ������
            For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                Try
                    tmpStartTime = CDate(FormatDateTime(dsTempTimingTask.Tables(0).Rows(i).Item("start_time"), DateFormat.ShortDate))
                    tmpProjectID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("project_code"))
                    tmpWorkFlowID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("workflow_id"))
                    tmpTaskID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("task_id"))
                    tmpMessageID = Trim(IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("message_id")), "", dsTempTimingTask.Tables(0).Rows(i).Item("message_id")))
                    tmpRoleID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("role_id"))
                    tmpStatus = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("Status"))
                    tmpDistance = IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("distance")), 0, dsTempTimingTask.Tables(0).Rows(i).Item("distance"))


                    '' �������ɫ��ȡ�붨ʱ�����PID���ID����ʾ��ɫƥ���Ա��������״̬��
                    'strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
                    'dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    Select Case Trim(dsTempTimingTask.Tables(0).Rows(i).Item("type"))

                        Case "T" '        �����ʱ��������Ϊ("T")

                            '        �������ʱ����Ŀ�ʼ����=��ǰ���ڣ�
                            If tmpStartTime <= sysTime Then
                                '   ��������ȡ�붨ʱ�����PID���IDƥ�������ID��
                                '   ����ʱ����״̬��Ϊ��E��
                                '   ����startupTask(ģ��ID��PID������ID)����ת������
                                '   ��������״̬�ǿգ�����Ŀ״̬��Ϊ�����ĵ�ǰ����״ֵ̬��
                                '   ��������ǰ����׶ηǿգ�����Ŀ�׶���Ϊ�����ĵ�ǰ����׶�ֵ��


                                '2010-8-3 yjf add �����ڱ�����Ƿ�Ϊ0��Ϊ�Ƿ����ڱ�׼
                                strSql = "select guaranting_sum from queryProjectInfoForStatistics where ProjectCode='" & tmpProjectID & "'"
                                dsTemp = CommonQuery.GetCommonQueryInfo(strSql)
                                If dsTemp.Tables(0).Rows.Count <> 0 Then

                                    If dsTemp.Tables(0).Rows(0).Item("guaranting_sum") <> 0 Then
                                        strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                                        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                                        dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


                                        WorkFlow.StartupTask(tmpWorkFlowID, tmpProjectID, tmpTaskID, "", "")

                                        strSql = "select staff_name from staff_role where isnull(overdue_message,0)=1"
                                        dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

                                        '2010-8-3 yjf add ����������Ϣ�����еķ�����
                                        For Each drTemp As DataRow In dsTemp.Tables(0).Rows
                                            AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, drTemp.Item("staff_name"), tmpStaffID, "9", "N")
                                        Next

                                        '����Ŀ��״̬��Ϊ�������Ӧ״̬
                                        tmpTaskPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))
                                        tmpTaskStatus = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_status")), "", dsTempTask.Tables(0).Rows(0).Item("project_status")))
                                        strSql = "{project_code=" & "'" & tmpProjectID & "'" & "}"
                                        dsTempProject = Project.GetProjectInfo(strSql)

                                        If tmpTaskPhase <> "" Then
                                            dsTempProject.Tables(0).Rows(0).Item("phase") = tmpTaskPhase
                                        End If

                                        If tmpTaskStatus <> "" Then
                                            dsTempProject.Tables(0).Rows(0).Item("status") = tmpTaskStatus
                                        End If


                                        Project.UpdateProject(dsTempProject)


                                        '����ʱ����״̬��Ϊ��E��[�û�ֻ������һ��]��
                                        dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                    End If

                                End If

                            End If


                        Case "P" '        �����ʱ��������Ϊ("P")
                            '���(��ʱ����Ŀ�ʼ���� = ��ǰ����)
                            '    ����ʱ����״̬��Ϊ��E����
                            '    �ڹ�����־��ȡ��������ɴ���HaveDone��
                            '    ��ȡ��ǰ��ʱ����״̬Ϊ��E�����������ShouldDone��
                            '    ��ȡ�ý�ɫ������Ա����
                            '       ��ÿ��Ա��()
                            '         ���ShouldDone>HaveDone
                            '         AddMsg����ĿID������ID����ϢID��Ա������N������

                            If tmpStartTime <= sysTime Then

                                dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                '��ȡ��ǰ��ʱ����״̬Ϊ��E�����������ShouldDone��
                                strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id='" & tmpRoleID & "' and type='P' and status='E'}"
                                dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                                tmpPromptCount = dsTemp.Tables(0).Rows.Count

                                'ͳ������ɴ���HaveDone��¼����
                                'strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                                'dsTempWorkLog = WorkLog.GetWorkLogInfo(strSql)
                                'tmpRecordCount = dsTempWorkLog.Tables(0).Rows.Count
                                strSql = "select COUNT(*) AS RecordCount from project_account_detail where  project_code = " & " '" & tmpProjectID & "'" & " and item_type='34' and item_code='001'"
                                tmpRecordCount = CommonQuery.GetCommonQueryInfo(strSql).Tables(0).Rows(0).Item("RecordCount")

                                ''��ȡ�ý�ɫ������Ա����
                                ''       ��ÿ��Ա��()
                                ''         ���ShouldDone>HaveDone
                                ''         AddMsg����ĿID������ID����ϢID��Ա������N������
                                'For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                                '    tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                                '    If tmpPromptCount > tmpRecordCount Then
                                '        AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N")
                                '    End If
                                'Next

                                '2010-05-13 YJF ADD 
                                If tmpPromptCount > tmpRecordCount Then
                                    AddMessgeToRealAccepterForRefund(tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpRoleID, tmpMessageID, tmpStartTime)
                                    ''����ʱ����״̬��Ϊ��E��[�û�ֻ������һ��]��
                                    '2010-05-24 YJF EDIT ����ʱ����״̬��Ϊ��P��[���δ��ʱ���������ʾ]��
                                    dsTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
                                    WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)
                                End If


                            End If

                        Case "W"
                                '�Զ��ָ�������������
                                If tmpStartTime <= sysTime Then
                                    WorkFlow.resumeProcess(tmpProjectID)
                                End If

                                ''����ʱ����״̬��Ϊ��E��[�û�ֻ������һ��]��
                                'dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                        Case "R"
                                'ѭ����������ʾ��Ϣ
                                '���(��ʱ����Ŀ�ʼ���� = ��ǰ����)
                                '���ö�ʱ����Ŀ�ʼʱ����Ϊ����Ŀ�ʼʱ��+����ļ��ʱ��
                                If tmpStartTime <= sysTime Then
                                    strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                                    dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                                    Dim tmpR As DateTime = IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("start_time")), Now, dsTempTimingTask.Tables(0).Rows(i).Item("start_time"))
                                    Dim tmpD As Integer = IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("distance")), 0, dsTempTimingTask.Tables(0).Rows(i).Item("distance"))
                                    dsTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Day, tmpD, tmpR)
                                    WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                End If

                                '2010-05-13 YJF ADD ���������ǲ�����Ŀ��������Ҫ����Ŀ������Ϊ�����˷�����Ϣ�����
                        Case "M"
                                If tmpStartTime <= sysTime Then

                                    '2010-05-13 YJF ADD 
                                    AddMessgeToRealAccepterForManagerA(tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpRoleID, tmpMessageID)

                                    '2010-5-12 yjf edit ��ѭ������(�������δ��ɣ����ָ��ʱ���������ѣ����򣬹رն�ʱ����)
                                    strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                                    dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                                    If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(0).Item("task_status")) = "P" Then
                                        dsTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Hour, tmpDistance * 24, tmpStartTime)
                                    Else
                                        dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                    End If
                                    WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                End If

                    End Select
                Catch errWf As WorkFlowErr
                    drTimeServerLog = dsTimeServerLog.Tables(0).NewRow()
                    drTimeServerLog.Item("time_server_log") = DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    tmpProjectID & " " & tmpTaskID & " " & errwf.ErrMessage
                    dsTimeServerLog.Tables(0).Rows.Add(drTimeServerLog)

                    'w.WriteLine(DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    'tmpProjectID & " " & tmpTaskID & " " & errwf.ErrMessage)
                Catch e As Exception
                    drTimeServerLog = dsTimeServerLog.Tables(0).NewRow()
                    drTimeServerLog.Item("time_server_log") = DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    tmpProjectID & " " & tmpTaskID & " " & e.Message
                    dsTimeServerLog.Tables(0).Rows.Add(drTimeServerLog)

                    'w.WriteLine(DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    'tmpProjectID & " " & tmpTaskID & " " & e.Message)
                End Try
            Next

        End If


        '�жϱ�Сʱ�Ƿ���Ҫɨ�裨��Գ�ʱ��Ϣ��
        If bHour = False Then

            '��	��ȡ��ʱ����״̬Ϊ��P��,��������Ϊ��A��������
            strSql = "{status='P' and type='A'}"
            dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

            '	����ÿһ������
            For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                Try

                    'tmpStartTime = CDate(FormatDateTime(dsTempTimingTask.Tables(0).Rows(i).Item("start_time"), DateFormat.ShortDate))
                    tmpStartTime = CDate(dsTempTimingTask.Tables(0).Rows(i).Item("start_time"))
                    tmpProjectID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("project_code"))
                    tmpWorkFlowID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("workflow_id"))
                    tmpTaskID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("task_id"))
                    tmpMessageID = Trim(IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("message_id")), "", dsTempTimingTask.Tables(0).Rows(i).Item("message_id")))
                    tmpRoleID = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("role_id"))
                    tmpStatus = Trim(dsTempTimingTask.Tables(0).Rows(i).Item("Status"))
                    tmpDistance = IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("distance")), 0, dsTempTimingTask.Tables(0).Rows(i).Item("distance"))

                    '�������ʱ����Ŀ�ʼʱ��=��ǰʱ�䣩(�����ں�Сʱͬʱ�ж�)
                    If (tmpStartTime < sysTime) Or (tmpStartTime = sysTime And Hour(tmpStartTime) <= Hour(Now)) Then

                        '' �������ɫ��ȡ�붨ʱ�����PID���ID����ʾ��ɫƥ���Ա��������״̬��
                        'strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
                        'dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


                        ''���Ա��������״̬��Ϊ�գ���ȡ��ʾ��ɫ��Ա��(�쵼)������Ϣ�������Ϣ����ĿID������ID�������ڡ���Ա��ID����
                        'If dsTempTaskAttenddee.Tables(0).Rows.Count = 0 Then

                        '    'tmpStaffID = WorkFlow.getTaskActor(tmpRoleID)

                        '    '2010-05-11 yjf edit �����Ż�ȡ��ɫ��Ա
                        '    strSql = "{project_code='" & tmpProjectID & "'}"
                        '    dsTempProject = Project.GetProjectInfo(strSql)
                        '    tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name"))
                        '    tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

                        '    If tmpStaffID = "" Then

                        '        '��ȡ��֧�������ϼ�����

                        '        strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                        '        dsBranch = Branch.GetBranch(strSql)

                        '        tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                        '        '��ȡ�ϼ������Ĳ�����ACTOR
                        '        tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

                        '    End If

                        '    '��ȡ�����������
                        '    strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                        '    dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                        '    '���쵼����ÿ�������˵���ʾ��Ϣ
                        '    If dsTempTaskAttenddee.Tables(0).Rows.Count <> 0 Then
                        '        For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                        '            tmpResponser = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                        '            AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N")
                        '        Next
                        '    Else
                        '        '���쵼����ÿ�������˵���ʾ��Ϣ
                        '        AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, "", tmpMessageID, "N")
                        '    End If
                        'Else
                        '    '        ����(�ǿ�)
                        '    '            ��ȡÿ�ý�ɫ״̬Ϊ��P����Ա����
                        '    '����Ϣ�������Ϣ����ĿID������ID�������ڡ���Ա��ID����

                        '    For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                        '        tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                        '        If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")) = "P" Then
                        '            AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N")
                        '        End If
                        '    Next
                        'End If


                        AddMessgeToRealAccepter(tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpRoleID, tmpMessageID)

                        ''����ʱ����״̬��Ϊ��E��[�û�ֻ������һ��]��
                        'dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                        'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                        '2010-5-12 yjf edit ��ѭ������(�������δ��ɣ����ָ��ʱ���������ѣ����򣬹رն�ʱ����)
                        strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                        dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                        If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(0).Item("task_status")) = "P" Then
                            dsTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Hour, tmpDistance * 24, sysTime)
                        Else
                            dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                        End If
                        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                    End If
                Catch errWf As WorkFlowErr
                    drTimeServerLog = dsTimeServerLog.Tables(0).NewRow()
                    drTimeServerLog.Item("time_server_log") = DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    tmpProjectID & " " & tmpTaskID & " " & errwf.ErrMessage
                    dsTimeServerLog.Tables(0).Rows.Add(drTimeServerLog)

                    'w.WriteLine(DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    'tmpProjectID & " " & tmpTaskID & " " & errwf.ErrMessage)
                Catch e As Exception
                    drTimeServerLog = dsTimeServerLog.Tables(0).NewRow()
                    drTimeServerLog.Item("time_server_log") = DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    tmpProjectID & " " & tmpTaskID & " " & e.Message
                    dsTimeServerLog.Tables(0).Rows.Add(drTimeServerLog)

                    'w.WriteLine(DateTime.Now.ToLongTimeString() & " " & DateTime.Now.ToLongDateString() & " " & _
                    'tmpProjectID & " " & tmpTaskID & " " & e.Message)
                End Try
            Next
        End If

        '���¶�ʱ����ֵ
        TimeServerLog.UpdateTimeServerLog(dsTimeServerLog)

        ''�ر���־�ļ�����
        'w.Close()
        'fs.Close()

    End Function

    Public Function AddMsg(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal tmpStaffID As String, ByVal messageID As String, ByVal readFlag As String)
        Dim strSql, tmpTaskName, tmpMsgType, msgContent As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet
        '��	��ȡ��ĿID�����ƣ�
        '��	��ȡ����ID���������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '��	��ȡ��ϢID����Ϣ���ƣ�
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '�쳣����
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '��ȡ��Ŀ����ҵ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '��	�����ϢID����Ϣ����Ϊ��Project�� ������ĿID+��Ϣ������Ϊ��ʾ��Ϣ��
        Select Case tmpMsgType
            Case "project"
                msgContent = tmpCorporationName & "��Ŀ" & msgContent
                '��	�����ϢID����Ϣ����Ϊ��Task�� ������ĿID+��������+��Ϣ������Ϊ��ʾ��Ϣ��
            Case "task"
                msgContent = tmpCorporationName & "��Ŀ��" & tmpTaskName & "����:" & msgContent
        End Select

        '��	����Ϣ�������Ϣ����ʾ��Ϣ��Ա������N������
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow

        With newRow
            .Item("project_code") = projectID
            .Item("message_content") = msgContent
            .Item("accepter") = tmpStaffID
            .Item("send_time") = Now
            .Item("is_affirmed") = readFlag
        End With
        dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
    End Function

    Public Function AddHeaderMsg(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal accepterID As String, ByVal responserID As String, ByVal messageID As String, ByVal readFlag As String)
        Dim strSql, tmpTaskName, tmpMsgType, msgContent As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet

        '��	��ȡ��ĿID�����ƣ�
        '��	��ȡ����ID���������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '��	��ȡ��ϢID����Ϣ���ƣ�
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '�쳣����
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '��ȡ��Ŀ����ҵ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '��	�����ϢID����Ϣ����Ϊ��Project�� ������ĿID+��Ϣ������Ϊ��ʾ��Ϣ��
        Select Case tmpMsgType
            Case "project"
                msgContent = responserID & " " & tmpCorporationName & "��Ŀ" & msgContent
                '��	�����ϢID����Ϣ����Ϊ��Task�� ������ĿID+��������+��Ϣ������Ϊ��ʾ��Ϣ��
            Case "task"
                msgContent = responserID & " " & tmpCorporationName & "��Ŀ��" & tmpTaskName & "����:" & msgContent
        End Select

        '��	����Ϣ�������Ϣ����ʾ��Ϣ��Ա������N������
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow

        With newRow
            .Item("project_code") = projectID
            .Item("message_content") = msgContent
            .Item("accepter") = accepterID
            .Item("responser") = responserID
            .Item("send_time") = Now
            .Item("is_affirmed") = readFlag
        End With
        dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
    End Function

    Public Function AddMsgContent(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal tmpStaffID As String, ByVal messageContent As String, ByVal readFlag As String)
        Dim strSql, tmpTaskName, tmpMsgType As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet
        '��	��ȡ��ĿID�����ƣ�
        '��	��ȡ����ID���������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")


        '��ȡ��Ŀ����ҵ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        messageContent = tmpCorporationName & "��Ŀ" & messageContent

        '��	����Ϣ�������Ϣ����ʾ��Ϣ��Ա������N������
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow

        With newRow
            .Item("project_code") = projectID
            .Item("message_content") = messageContent
            .Item("accepter") = tmpStaffID
            .Item("send_time") = Now
            .Item("is_affirmed") = readFlag
        End With
        dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
    End Function


    Private Function AddMessgeToRealAccepter(ByVal tmpProjectID As String, ByVal tmpWorkFlowID As String, ByVal tmpTaskID As String, ByVal tmpRoleID As String, ByVal tmpMessageID As String) As String
        Dim strSql, tmpBranch, tmpStaffID, tmpSuper, tmpResponser As String
        Dim dsTempTaskAttenddee, dsTempProject, dsBranch, dsTemp As DataSet

        Dim j As Integer

        ' �������ɫ��ȡ�붨ʱ�����PID���ID����ʾ��ɫƥ���Ա��������״̬��
        strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
        dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '2010-8-3 yjf add ���Ϊ�Ǽǻ���֤��������,���ͳ�ʱ��Ϣ�����з�����
        If tmpTaskID = "RecordRefundCertificate" And tmpRoleID = "33" Then
            strSql = "select staff_name from staff_role where isnull(overdue_message,0)=1"
            dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

            '2010-8-3 yjf add ����������Ϣ�����еķ�����
            For Each drTemp As DataRow In dsTemp.Tables(0).Rows
                AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, drTemp.Item("staff_name"), tmpStaffID, tmpMessageID, "N")
            Next

            Exit Function
        End If


        '���Ա��������״̬��Ϊ�գ���ȡ��ʾ��ɫ��Ա��(�쵼)������Ϣ�������Ϣ����ĿID������ID�������ڡ���Ա��ID����
        If dsTempTaskAttenddee.Tables(0).Rows.Count = 0 Then

            'tmpStaffID = WorkFlow.getTaskActor(tmpRoleID)

            '2010-05-11 yjf edit �����Ż�ȡ��ɫ��Ա
            strSql = "select deptName from queryProjectInfo where ProjectCode='" & tmpProjectID & "'"
            'dsTempProject = Project.GetProjectInfo(strSql)
            dsTempProject = CommonQuery.GetCommonQueryInfo(strSql)
            tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("deptName")), "", dsTempProject.Tables(0).Rows(0).Item("deptName"))
            tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

            If tmpStaffID = "" Then

                '��ȡ��֧�������ϼ�����

                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                dsBranch = Branch.GetBranch(strSql)

                tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '��ȡ�ϼ������Ĳ�����ACTOR
                tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

            End If

            '��ȡ�����������
            strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '���쵼����ÿ�������˵���ʾ��Ϣ
            If dsTempTaskAttenddee.Tables(0).Rows.Count <> 0 Then
                For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                    tmpResponser = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                    AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N")
                Next
            Else
                '���쵼����ÿ�������˵���ʾ��Ϣ
                AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, "", tmpMessageID, "N")
            End If
        Else
            '        ����(�ǿ�)
            '            ��ȡÿ�ý�ɫ״̬Ϊ��P����Ա����
            '����Ϣ�������Ϣ����ĿID������ID�������ڡ���Ա��ID����

            For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")) = "P" Then
                    AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N")
                End If
            Next
        End If

    End Function


    Private Function AddMessgeToRealAccepterForManagerA(ByVal tmpProjectID As String, ByVal tmpWorkFlowID As String, ByVal tmpTaskID As String, ByVal tmpRoleID As String, ByVal tmpMessageID As String) As String
        Dim strSql, tmpBranch, tmpStaffID, tmpSuper, tmpResponser As String
        Dim dsTempTaskAttenddee, dsTempProject, dsBranch As DataSet

        Dim j As Integer

        strSql = "select top 1 attend_person from project_task_attendee where project_code='" & tmpProjectID & "' and role_id='24'"
        Dim objDsCommonQuery As DataSet = CommonQuery.GetCommonQueryInfo(strSql)
        tmpResponser = objDsCommonQuery.Tables(0).Rows(0).Item("attend_person")

        '�����ɫ��Ϊ��Ŀ����,�����쵼������Ϣ
        If tmpRoleID <> "24" Then

            '2010-05-11 yjf edit �����Ż�ȡ��ɫ��Ա
            strSql = "{project_code='" & tmpProjectID & "'}"
            dsTempProject = Project.GetProjectInfo(strSql)
            tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name"))
            tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

            If tmpStaffID = "" Then

                '��ȡ��֧�������ϼ�����

                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                dsBranch = Branch.GetBranch(strSql)

                tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '��ȡ�ϼ������Ĳ�����ACTOR
                tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

            End If

            AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N")

        Else
            '��������Ŀ�����˷�����Ϣ
            AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpResponser, tmpMessageID, "N")

        End If

    End Function


    Public Function AddHeaderMsgForRefund(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal accepterID As String, ByVal responserID As String, ByVal messageID As String, ByVal readFlag As String, ByVal tmpStartTime As Date)
        Dim strSql, tmpTaskName, tmpMsgType, msgContent As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet

        '��	��ȡ��ĿID�����ƣ�
        '��	��ȡ����ID���������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '��	��ȡ��ϢID����Ϣ���ƣ�
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '�쳣����
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '��ȡ��Ŀ����ҵ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '��	�����ϢID����Ϣ����Ϊ��Project�� ������ĿID+��Ϣ������Ϊ��ʾ��Ϣ��
        Select Case tmpMsgType
            Case "project"
                msgContent = responserID & " " & tmpCorporationName & "��Ŀ" & " " & tmpStartTime.ToShortDateString & "Ӧ" & msgContent
                '��	�����ϢID����Ϣ����Ϊ��Task�� ������ĿID+��������+��Ϣ������Ϊ��ʾ��Ϣ��
            Case "task"
                msgContent = responserID & " " & tmpCorporationName & "��Ŀ��" & tmpTaskName & "����:" & msgContent
        End Select

        '��	����Ϣ�������Ϣ����ʾ��Ϣ��Ա������N������
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow

        With newRow
            .Item("project_code") = projectID
            .Item("message_content") = msgContent
            .Item("accepter") = accepterID
            .Item("responser") = responserID
            .Item("send_time") = Now
            .Item("is_affirmed") = readFlag
        End With
        dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
    End Function

    Public Function AddMsgForRefund(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal tmpStaffID As String, ByVal messageID As String, ByVal readFlag As String, ByVal tmpStartTime As Date)
        Dim strSql, tmpTaskName, tmpMsgType, msgContent As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet
        '��	��ȡ��ĿID�����ƣ�
        '��	��ȡ����ID���������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '��	��ȡ��ϢID����Ϣ���ƣ�
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '�쳣����
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '��ȡ��Ŀ����ҵ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '��	�����ϢID����Ϣ����Ϊ��Project�� ������ĿID+��Ϣ������Ϊ��ʾ��Ϣ��
        Select Case tmpMsgType
            Case "project"
                msgContent = tmpCorporationName & "��Ŀ" & " " & tmpStartTime.ToShortDateString & "Ӧ" & msgContent
                '��	�����ϢID����Ϣ����Ϊ��Task�� ������ĿID+��������+��Ϣ������Ϊ��ʾ��Ϣ��
            Case "task"
                msgContent = tmpCorporationName & "��Ŀ��" & tmpTaskName & "����:" & msgContent
        End Select

        '��	����Ϣ�������Ϣ����ʾ��Ϣ��Ա������N������
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow

        With newRow
            .Item("project_code") = projectID
            .Item("message_content") = msgContent
            .Item("accepter") = tmpStaffID
            .Item("send_time") = Now
            .Item("is_affirmed") = readFlag
        End With
        dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)
    End Function

    Private Function AddMessgeToRealAccepterForRefund(ByVal tmpProjectID As String, ByVal tmpWorkFlowID As String, ByVal tmpTaskID As String, ByVal tmpRoleID As String, ByVal tmpMessageID As String, ByVal tmpStartTime As Date) As String
        Dim strSql, tmpBranch, tmpStaffID, tmpSuper, tmpResponser As String
        Dim dsTempTaskAttenddee, dsTempProject, dsBranch As DataSet

        Dim j As Integer

        ' �������ɫ��ȡ�붨ʱ�����PID���ID����ʾ��ɫƥ���Ա��������״̬��
        strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
        dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        '���Ա��������״̬��Ϊ�գ���ȡ��ʾ��ɫ��Ա��(�쵼)������Ϣ�������Ϣ����ĿID������ID�������ڡ���Ա��ID����
        If dsTempTaskAttenddee.Tables(0).Rows.Count = 0 Then

            'tmpStaffID = WorkFlow.getTaskActor(tmpRoleID)

            '2010-05-11 yjf edit �����Ż�ȡ��ɫ��Ա
            strSql = "{project_code='" & tmpProjectID & "'}"
            dsTempProject = Project.GetProjectInfo(strSql)
            tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name"))
            tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

            If tmpStaffID = "" Then

                '��ȡ��֧�������ϼ�����

                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                dsBranch = Branch.GetBranch(strSql)

                tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '��ȡ�ϼ������Ĳ�����ACTOR
                tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

            End If

            '��ȡ�����������
            strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '���쵼����ÿ�������˵���ʾ��Ϣ
            If dsTempTaskAttenddee.Tables(0).Rows.Count <> 0 Then
                For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                    tmpResponser = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                    AddHeaderMsgForRefund(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N", tmpStartTime)
                Next
            Else
                '���쵼����ÿ�������˵���ʾ��Ϣ
                AddHeaderMsgForRefund(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, "", tmpMessageID, "N", tmpStartTime)
            End If
        Else
            '        ����(�ǿ�)
            '            ��ȡÿ�ý�ɫ״̬Ϊ��P����Ա����
            '����Ϣ�������Ϣ����ĿID������ID�������ڡ���Ա��ID����

            For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")) = "P" Then
                    AddMsgForRefund(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N", tmpStartTime)
                End If
            Next
        End If

    End Function

End Class
