Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WorkFlow

    '����ģ�峣��
    Public Const Table_Task_Template As String = "task_template"
    Public Const Table_Task_Transfer_Template As String = "task_transfer_template"
    Public Const Table_Task_Role_Template As String = "task_role_template"
    Public Const Table_Timing_Task_Template As String = "timing_task_template"


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfTemplate As SqlDataAdapter

    '�����ѯ����
    Private GetWfTemplateInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���幤������������
    Private WfProjectTask As WfProjectTask
    Private WfProjectMessages As WfProjectMessages
    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTimingTask As WfProjectTimingTask
    Private WfProjectTrack As WfProjectTrack
    Private ProjectResponsible As ProjectResponsible


    '������Ŀ����
    Private project As project

    '���幤����־��������
    Private WorkLog As WorkLog

    Private WorkflowType As WorkflowType

    Private commQuery As CommonQuery

    Private staff As staff

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        dsCommand_WfTemplate = New SqlDataAdapter()

        '�����ⲿ����
        ts = trans

        'ʵ��������������
        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectMessages = New WfProjectMessages(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        WfProjectTrack = New WfProjectTrack(conn, ts)

        ProjectResponsible = New ProjectResponsible(conn, ts)

        'ʵ������Ŀ����
        project = New Project(conn, ts)

        'ʵ����������־����
        WorkLog = New WorkLog(conn, ts)

        WorkflowType = New WorkflowType(conn, ts)

        commQuery = New CommonQuery(conn, ts)

        staff = New Staff(conn, ts)

    End Sub

    '��ȡ������ģ����Ϣ
    Public Function GetWfProjectTaskTemplateInfo(ByVal templateTableName As String, ByVal condition As String) As DataSet

        Dim tempDs As New DataSet()

        Dim strSql As String = "select * from " & templateTableName & " where " & condition
        GetWfTemplateInfoCommand = New SqlCommand(strSql, conn)
        GetWfTemplateInfoCommand.CommandType = CommandType.Text

        With dsCommand_WfTemplate
            .SelectCommand = GetWfTemplateInfoCommand
            .SelectCommand.Transaction = ts
            .Fill(tempDs, Table_Task_Template)
        End With

        Return tempDs
    End Function

    '����Ϣ���з�����Ϣ
    Private Function SendMessage()

    End Function

    '��������
    '2005-03-18 yjf add ����ǰ������ǰ������ID��ǰ������������
    Public Function StartupTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal rollbackMsg As String, ByVal responserID As String, ByVal submitTaskID As String, ByVal submitUser As String)

        '����ԭ�������񷽷�
        StartupTask(workFlowID, projectID, taskID, rollbackMsg, responserID)

        Dim strSql As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
            '2005-03-18 yjf add ����ǰ������ǰ������ID��ǰ������������
            dsTempTaskStatus.Tables(0).Rows(i).Item("previous_task_id") = submitTaskID
            dsTempTaskStatus.Tables(0).Rows(i).Item("previous_task_attendee") = submitUser
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)
    End Function

    '��������
    Public Function StartupTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal rollbackMsg As String, ByVal responserID As String)

        '��������ȡ�����(ģ��ID����ĿID������ID)ƥ�������
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '��ȡϵͳʱ�䣻�����������Ŀ�ʼʱ���Ϊϵͳʱ��
        Dim sysTime As DateTime = Now

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        dsTempTask.Tables(0).Rows(0).Item("start_time") = sysTime

        '����������������ģʽΪ"Manual",�����ÿ�
        Dim tmpStartMode As String = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_mode")), "", dsTempTask.Tables(0).Rows(0).Item("start_mode")))
        If tmpStartMode = "manual" Then
            dsTempTask.Tables(0).Rows(0).Item("start_mode") = DBNull.Value
        End If

        WfProjectTask.UpdateWfProjectTask(dsTempTask)


        '��ȡ�����Ƿ��跢����Ϣ
        Dim bTmp As Boolean = False
        If Not dsTempTask.Tables(0).Rows(0).Item("hasMessage") Is DBNull.Value Then
            bTmp = dsTempTask.Tables(0).Rows(0).Item("hasMessage")
        End If

        Dim tmpAttend, tmpBranch As String


        '������������Ϊ��(�Ƿ����ɫ)(��ί��Ȩ�޵Ľ�ɫ����ֻ��һ�������¼)


        '  ����getTaskActor��RoleID����ȡ��������ˣ�
        '  ����ǰ����Ĳ�������Ϊ��ȡ����������ˣ�

        Dim dsTempTaskAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣���� 
        If dsTempTaskAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttend.Tables(0))
            Throw wfErr
        End If

        'tmpAttend = IIf(IsDBNull(dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person")), "", dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person"))
        'Dim tmpRoleID As String = dsTempTaskAttend.Tables(0).Rows(0).Item("role_id")
        'Dim staff As New Staff(conn, ts)


        ''���ͻ�����Ϣ
        'If rollbackMsg <> "" Then
        '    AddMsg(projectID, taskID, rollbackMsg, tmpAttend, responserID)
        'End If

        'If tmpAttend = "" Then

        '    '��ȡ��ʼ����Ĳ����˵ķ�֧������
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and  task_type='BEGIN'}"
        '    Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '    '�쳣����  
        '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '        Dim wfErr As New WorkFlowErr()
        '        wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '        Throw wfErr
        '    End If

        '    Dim tmpBeginTask As String = dsTemp.Tables(0).Rows(0).Item("task_id")
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpBeginTask & "'" & "}"
        '    dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    '�쳣����  
        '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '        Dim wfErr As New WorkFlowErr()
        '        wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '        Throw wfErr
        '    End If

        '    tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
        '    strSql = "{staff_name=" & "'" & tmpAttend & "'" & "}"
        '    dsTemp = staff.FetchStaff(strSql)

        '    '�쳣����  
        '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '        Dim wfErr As New WorkFlowErr()
        '        wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '        Throw wfErr
        '    End If


        '    tmpBranch = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("branch_name")), "", dsTemp.Tables(0).Rows(0).Item("branch_name"))

        '    '����getTaskActor��RoleID����֧��������ȡ���������ACTOR��
        '    tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpBranch)

        '    '���ACTORΪ�գ�����getTaskActor��RoleID����ȡ���������ACTOR��
        '    If tmpAttend = "" Then

        '        '��ȡ��֧�������ϼ�����
        '        Dim Branch As New Branch(conn, ts)
        '        strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
        '        Dim dsBranch As DataSet = Branch.GetBranch(strSql)

        '        '�쳣����  
        '        If dsBranch.Tables(0).Rows.Count = 0 Then
        '            Dim wfErr As New WorkFlowErr()
        '            wfErr.ThrowNoRecordkErr(dsBranch.Tables(0))
        '            Throw wfErr
        '        End If

        '        Dim tmpSuper As String = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

        '        '��ȡ�ϼ������Ĳ�����ACTOR
        '        tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpSuper)

        '        'tmpAttend = getTaskActor(tmpRoleID)
        '    End If

        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        '    dsTempTaskAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person") = tmpAttend
        '    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttend)
        'End If

        'qxd modify 2004-10-11
        'Ϊ�˽��һ��������������ɫ�������⡣
        Dim k, count As Integer

        count = dsTempTaskAttend.Tables(0).Rows.Count

        Dim dsTemp As DataSet

        For k = 0 To count - 1
            tmpAttend = IIf(IsDBNull(dsTempTaskAttend.Tables(0).Rows(k).Item("attend_person")), "", dsTempTaskAttend.Tables(0).Rows(k).Item("attend_person"))
            Dim tmpRoleID As String = dsTempTaskAttend.Tables(0).Rows(k).Item("role_id")

            '���ͻ�����Ϣ
            If rollbackMsg <> "" Then
                AddMsg(projectID, taskID, rollbackMsg, tmpAttend, responserID)
            End If

            If tmpAttend = "" Then

                ''��ȡ��ʼ����Ĳ����˵ķ�֧������
                'strSql = "{project_code=" & "'" & projectID & "'" & " and  task_type='BEGIN'}"
                'Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

                ''�쳣����  
                'If dsTemp.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                '    Throw wfErr
                'End If

                'Dim tmpBeginTask As String = dsTemp.Tables(0).Rows(0).Item("task_id")
                'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpBeginTask & "'" & "}"
                'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                ''�쳣����  
                'If dsTemp.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                '    Throw wfErr
                'End If

                'tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
                'strSql = "{staff_name=" & "'" & tmpAttend & "'" & "}"
                'dsTemp = staff.FetchStaff(strSql)

                ''�쳣����  
                'If dsTemp.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                '    Throw wfErr
                'End If

                strSql = "{project_code='" & projectID & "'}"
                dsTemp = project.GetProjectInfo(strSql)

                '�쳣����  
                If dsTemp.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr()
                    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                    Throw wfErr
                End If

                tmpBranch = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTemp.Tables(0).Rows(0).Item("applicantTeam_name"))

                '����getTaskActor��RoleID����֧��������ȡ���������ACTOR��
                tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpBranch)

                '���ACTORΪ�գ�����getTaskActor��RoleID����ȡ���������ACTOR��
                If tmpAttend = "" Then

                    '��ȡ��֧�������ϼ�����
                    Dim Branch As New Branch(conn, ts)
                    strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                    Dim dsBranch As DataSet = Branch.GetBranch(strSql)

                    '�쳣����  
                    If dsBranch.Tables(0).Rows.Count = 0 Then
                        Dim wfErr As New WorkFlowErr()
                        wfErr.ThrowNoRecordkErr(dsBranch.Tables(0))
                        Throw wfErr
                    End If

                    Dim tmpSuper As String = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                    '��ȡ�ϼ������Ĳ�����ACTOR
                    tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpSuper)

                    'tmpAttend = getTaskActor(tmpRoleID)
                End If

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
                dsTempTaskAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                dsTempTaskAttend.Tables(0).Rows(k).Item("attend_person") = tmpAttend
                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttend)


            End If

            '2007-10-16 yjf edit ������Ϣ������ˢ�±��λ
            If tmpAttend <> "" Then

                strSql = "{staff_name='" & tmpAttend & "'}"
                dsTemp = staff.FetchStaff(strSql)

                '�쳣����  
                If dsTemp.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                    Throw wfErr
                End If

                If IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("DoScan")), False, dsTemp.Tables(0).Rows(0).Item("DoScan")) = False Then
                    dsTemp.Tables(0).Rows(0).Item("DoScan") = 1
                    staff.UpdateStaff(dsTemp)
                End If

            End If
        Next
        'qxd modify 2004-10-11 end 


        '��ÿ�������˵�ת������״̬��Ϊ��P��
        Dim TimingServer As New TimingServer(conn, ts, True, True)
        Dim tmpUserID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
            tmpUserID = Trim(dsTempTaskStatus.Tables(0).Rows(i).Item("attend_person"))
            dsTempTaskStatus.Tables(0).Rows(i).Item("task_status") = "P"
            '������Ϣ
            If bTmp Then
                TimingServer.AddMsg(workFlowID, projectID, taskID, tmpUserID, "16", "N")
            End If
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)

        '�ڶ�ʱ������뵱ǰ����IDƥ��Ķ�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and type='A'}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)


        Dim tmpTimeLimit As Integer

        '�����ж�ʱ����Ŀ�ʼʱ����Ϊ����Ŀ�ʼʱ��+��ʾʱ��
        '�����ж�ʱ������״̬��Ϊ��P����
        Dim newRow As DataRow
        '2010-8-3 yjf add ������Ϣ���Ǽǻ���֤������Ϣ���⣨��Ϊ�������������Ϣ��ImplAddRefundPlan�ӿ���ӽ�ȥ�Ժ󣬻������������������ʱ��Ὣ��Ԥ����Ϣ������ʱ���Ƴ٣���Ϊ������ʱ�䣽��ǰʱ�䣫��ʾ�����
        If taskID <> "OverdueRecord" And taskID <> "RecordRefundCertificate" Then

            For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                '��ʾ����
                tmpTimeLimit = dsTempTimingTask.Tables(0).Rows(i).Item("time_limit")
                newRow = dsTempTimingTask.Tables(0).Rows(i)
                With newRow
                    .Item("start_time") = DateAdd(DateInterval.Hour, tmpTimeLimit * 24, sysTime) '����ʱ�䣽��ǰʱ�䣫��ʾ���
                    .Item("status") = "P"
                End With
            Next

        End If

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


    End Function

    '��������
    Public Function StartupManualTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal rollbackMsg As String, ByVal responserID As String)

        '��������ȡ�����(ģ��ID����ĿID������ID)ƥ�������
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '��ȡϵͳʱ�䣻�����������Ŀ�ʼʱ���Ϊϵͳʱ��
        Dim sysTime As DateTime = Now

        '�쳣����
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        dsTempTask.Tables(0).Rows(0).Item("start_time") = sysTime
        WfProjectTask.UpdateWfProjectTask(dsTempTask)


        '��ȡ�����Ƿ��跢����Ϣ
        Dim bTmp As Boolean = False
        If Not dsTempTask.Tables(0).Rows(0).Item("hasMessage") Is DBNull.Value Then
            bTmp = dsTempTask.Tables(0).Rows(0).Item("hasMessage")
        End If

        Dim tmpAttend, tmpBranch As String


        '������������Ϊ��(�Ƿ����ɫ)(��ί��Ȩ�޵Ľ�ɫ����ֻ��һ�������¼)


        '  ����getTaskActor��RoleID����ȡ��������ˣ�
        '  ����ǰ����Ĳ�������Ϊ��ȡ����������ˣ�

        Dim dsTempTaskAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣���� 
        If dsTempTaskAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttend.Tables(0))
            Throw wfErr
        End If

        tmpAttend = IIf(IsDBNull(dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person")), "", dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person"))
        Dim tmpRoleID As String = dsTempTaskAttend.Tables(0).Rows(0).Item("role_id")
        Dim staff As New Staff(conn, ts)


        '���ͻ�����Ϣ
        If rollbackMsg <> "" Then
            AddMsg(projectID, taskID, rollbackMsg, tmpAttend, responserID)
        End If

        If tmpAttend = "" Then

            '��ȡ��ʼ����Ĳ����˵ķ�֧������
            strSql = "{project_code=" & "'" & projectID & "'" & " and  task_type='BEGIN'}"
            Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            Dim tmpBeginTask As String = dsTemp.Tables(0).Rows(0).Item("task_id")
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpBeginTask & "'" & "}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
            strSql = "{staff_name=" & "'" & tmpAttend & "'" & "}"
            dsTemp = staff.FetchStaff(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If


            tmpBranch = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("branch_name")), "", dsTemp.Tables(0).Rows(0).Item("branch_name"))

            '����getTaskActor��RoleID����֧��������ȡ���������ACTOR��
            tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpBranch)

            '���ACTORΪ�գ�����getTaskActor��RoleID����ȡ���������ACTOR��
            If tmpAttend = "" Then

                '��ȡ��֧�������ϼ�����
                Dim Branch As New Branch(conn, ts)
                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                Dim dsBranch As DataSet = Branch.GetBranch(strSql)

                '�쳣����  
                If dsBranch.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr()
                    wfErr.ThrowNoRecordkErr(dsBranch.Tables(0))
                    Throw wfErr
                End If

                Dim tmpSuper As String = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '��ȡ�ϼ������Ĳ�����ACTOR
                tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpSuper)

                'tmpAttend = getTaskActor(tmpRoleID)
            End If

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
            dsTempTaskAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person") = tmpAttend
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttend)
        End If


        '��ÿ�������˵�ת������״̬��Ϊ��P��
        Dim TimingServer As New TimingServer(conn, ts, True, True)
        Dim tmpUserID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
            tmpUserID = Trim(dsTempTaskStatus.Tables(0).Rows(i).Item("attend_person"))
            dsTempTaskStatus.Tables(0).Rows(i).Item("task_status") = "P"
            '������Ϣ
            If bTmp Then
                TimingServer.AddMsg(workFlowID, projectID, taskID, tmpUserID, "16", "N")
            End If
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)

        '�ڶ�ʱ������뵱ǰ����IDƥ��Ķ�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and type='A'}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)


        Dim tmpTimeLimit As Integer

        '�����ж�ʱ����Ŀ�ʼʱ����Ϊ����Ŀ�ʼʱ��+��ʾʱ��
        '�����ж�ʱ������״̬��Ϊ��P����
        Dim newRow As DataRow
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            '��ʾ����
            tmpTimeLimit = dsTempTimingTask.Tables(0).Rows(i).Item("time_limit")
            newRow = dsTempTimingTask.Tables(0).Rows(i)
            With newRow
                .Item("start_time") = DateAdd(DateInterval.Hour, tmpTimeLimit * 24, sysTime) '����ʱ�䣽��ǰʱ�䣫��ʾ���
                .Item("status") = "P"
            End With
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


    End Function

    '����������
    Public Function CreateProcess(ByVal workFlowID As String, ByVal projectID As String, ByVal userID As String) As String
        CreateProcess(workFlowID, projectID, userID, "1")
    End Function

    '����������
    Public Function CreateProcess(ByVal workFlowID As String, ByVal projectID As String, ByVal userID As String, ByVal phase As String) As String


        '��ȡϵͳʱ��
        Dim sysTime As DateTime = Today

        Dim strSql As String

        ''��ȡ��Ŀ�׶�
        Dim tmpTaskPhase, tmpTaskStatus As String
        Dim dsTempProject As DataSet
        If phase = "1" Then
            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)
            '�쳣����  
            If dsTempProject.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                Throw wfErr
            End If

            tmpTaskPhase = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("phase")), "", dsTempProject.Tables(0).Rows(0).Item("phase"))
        Else
            tmpTaskPhase = phase
        End If


        '����ҵ��Ʒ�ֺ���Ŀ�׶λ�ȡģ��ID
        strSql = "{service_type=" & "'" & workFlowID & "'" & " and isnull(phase,'')=" & "'" & tmpTaskPhase & "'" & "}"
        Dim dsWorkflowType As DataSet = WorkflowType.GetWorkflowTypeInfo(strSql)

        '�쳣����  
        If dsWorkflowType.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsWorkflowType.Tables(0))
            Throw wfErr
        End If

        Dim strWorkflowID As String = dsWorkflowType.Tables(0).Rows(0).Item("workflow_id")
        Dim strWorkflow As String = "workflow_id=" & "'" & strWorkflowID & "'"
        'Dim strWorkflow As String = "workflow_id='01'"

        '�����������Ƿ���ڷ���������ָ���Ĺ���������
        '��������ڣ��������������󣬷����쳣����
        strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id=" & "'" & strWorkflowID & "'" & "}"
        Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)
        If dsTemp.Tables(0).Rows.Count <> 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowExistWorkFlowErr()
            Throw wfErr
        Else

            '1������ģ��
            Dim dsTemplate As DataSet = GetWfProjectTaskTemplateInfo("task_template", strWorkflow)

            Dim newRow As DataRow
            Dim i As Integer
            Dim straTime As DateTime = Now
            Dim beginTaskID As String

            '������ҵ����������ģ�����������ӵ������
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1


                newRow = dsTemp.Tables(0).NewRow()
                With newRow

                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                    .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                    .Item("sequence") = dsTemplate.Tables(0).Rows(i).Item("sequence")
                    .Item("task_name") = dsTemplate.Tables(0).Rows(i).Item("task_name")
                    .Item("task_type") = dsTemplate.Tables(0).Rows(i).Item("task_type")
                    .Item("apply_tool") = dsTemplate.Tables(0).Rows(i).Item("apply_tool")
                    .Item("parameters") = dsTemplate.Tables(0).Rows(i).Item("parameters")
                    .Item("duration") = dsTemplate.Tables(0).Rows(i).Item("duration")
                    .Item("merge_relation") = dsTemplate.Tables(0).Rows(i).Item("merge_relation")
                    .Item("flow_tool") = dsTemplate.Tables(0).Rows(i).Item("flow_tool")
                    .Item("create_person") = dsTemplate.Tables(0).Rows(i).Item("create_person")
                    .Item("create_date") = dsTemplate.Tables(0).Rows(i).Item("create_date")
                    .Item("project_phase") = dsTemplate.Tables(0).Rows(i).Item("phase")
                    .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                    .Item("hasMessage") = dsTemplate.Tables(0).Rows(i).Item("hasMessage")

                End With

                '��ȡ����������ʼ�����
                If IIf(IsDBNull(dsTemplate.Tables(0).Rows(i).Item("task_type")), "", dsTemplate.Tables(0).Rows(i).Item("task_type")) = "BEGIN" Then
                    beginTaskID = Trim(dsTemplate.Tables(0).Rows(i).Item("task_id"))
                    newRow.Item("start_time") = Now
                End If
                dsTemp.Tables(0).Rows.Add(newRow)
            Next

            WfProjectTask.UpdateWfProjectTask(dsTemp)

            '2����ɫģ��
            dsTemplate = GetWfProjectTaskTemplateInfo("task_role_template", strWorkflow)
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo("null")

            '����ɫģ����ӵ���ɫ����
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
                newRow = dsTemp.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                    .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                    .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")

                    '����������ʼ������Ա��ID������Ϊ��ѯ��ԱID,������������״̬��Ϊ��P�������У�
                    If Trim(dsTemplate.Tables(0).Rows(i).Item("task_id")) = beginTaskID Then
                        .Item("attend_person") = userID
                        .Item("task_status") = "P"
                    Else
                        .Item("attend_person") = ""
                    End If

                End With
                dsTemp.Tables(0).Rows.Add(newRow)
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            '3��ת������ģ��
            dsTemplate = GetWfProjectTaskTemplateInfo("task_transfer_template", strWorkflow)
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")

            '��ת������ģ����ӵ�ת����������
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
                newRow = dsTemp.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                    .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                    .Item("next_task") = dsTemplate.Tables(0).Rows(i).Item("next_task")
                    .Item("transfer_condition") = dsTemplate.Tables(0).Rows(i).Item("transfer_condition")
                    .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                    .Item("isItem") = dsTemplate.Tables(0).Rows(i).Item("isItem")
                End With
                dsTemp.Tables(0).Rows.Add(newRow)
            Next

            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)


            '4����ʱ����ģ��
            dsTemplate = GetWfProjectTaskTemplateInfo("timing_task_template", strWorkflow)
            dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")

            '������ģ����ӵ�����ģ��ʵ������
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
                newRow = dsTemp.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                    .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                    .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")
                    .Item("distance") = dsTemplate.Tables(0).Rows(i).Item("distance")
                    .Item("start_time") = "1900-01-01"
                    .Item("message_id") = dsTemplate.Tables(0).Rows(i).Item("message_id")
                    .Item("type") = dsTemplate.Tables(0).Rows(i).Item("type")
                    .Item("time_limit") = dsTemplate.Tables(0).Rows(i).Item("time_limit")
                    .Item("parameter") = dsTemplate.Tables(0).Rows(i).Item("parameter")
                    .Item("create_person") = dsTemplate.Tables(0).Rows(i).Item("create_person")
                    .Item("create_date") = dsTemplate.Tables(0).Rows(i).Item("create_date")
                End With
                dsTemp.Tables(0).Rows.Add(newRow)
            Next

            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)


            '	�����ǰ�������Ŀ�׶ηǿգ�����Ŀ�׶���Ϊ��ǰ����Ľ׶�ֵ��
            Dim dsTempTask As DataSet

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & beginTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '�쳣����  
            If dsTempTask.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                Throw wfErr
            End If

            tmpTaskPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))
            tmpTaskStatus = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_status")), "", dsTempTask.Tables(0).Rows(0).Item("project_status")))


            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)

            '�쳣����  
            If dsTempProject.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                Throw wfErr
            End If

            If tmpTaskPhase <> "" Then
                dsTempProject.Tables(0).Rows(0).Item("phase") = tmpTaskPhase
            End If

            '	�����ǰ�������Ŀ״̬�ǿգ�����Ŀ״̬��Ϊ��ǰ�����״̬

            If tmpTaskStatus <> "" Then
                dsTempProject.Tables(0).Rows(0).Item("status") = tmpTaskStatus
            End If

            project.UpdateProject(dsTempProject)


            finishedTask(strWorkflowID, projectID, beginTaskID, ".T.", userID)



        End If

    End Function

    '����ǩԼ����
    Public Function CopySignature(ByVal fatherProject As String, ByVal sonProject As String)
        Dim dsTempFather, dsTempSon As DataSet
        Dim strSql As String
        Dim objProjectSignature As New ProjectSignature(conn, ts)
        strSql = "{project_code='" & fatherProject & "'}"
        dsTempFather = objProjectSignature.GetProjectSignatureInfo(strSql)
        strSql = "{project_code='" & sonProject & "'}"
        dsTempSon = objProjectSignature.GetProjectSignatureInfo(strSql)
        Dim fatherRow As DataRow = dsTempFather.Tables(0).Rows(0)
        If dsTempSon.Tables(0).Rows.Count = 0 Then
            Dim sonRow As DataRow = dsTempSon.Tables(0).NewRow()
            With sonRow
                .Item("project_code") = sonProject
                .Item("sign_date") = fatherRow.Item("sign_date")
                .Item("remark") = fatherRow.Item("remark")
                .Item("create_person") = fatherRow.Item("create_person")
                .Item("create_date") = fatherRow.Item("create_date")
                .Item("sign_sum") = fatherRow.Item("sign_sum")
                .Item("bank") = fatherRow.Item("bank")
                .Item("bank_branch") = fatherRow.Item("bank_branch")
                .Item("loanContract_num") = fatherRow.Item("loanContract_num")
                .Item("assureContract_num") = fatherRow.Item("assureContract_num")
            End With
            dsTempSon.Tables(0).Rows.Add(sonRow)
        Else
            Dim sonRow As DataRow = dsTempSon.Tables(0).Rows(0)
            With sonRow
                .Item("project_code") = sonProject
                .Item("sign_date") = fatherRow.Item("sign_date")
                .Item("remark") = fatherRow.Item("remark")
                .Item("create_person") = fatherRow.Item("create_person")
                .Item("create_date") = fatherRow.Item("create_date")
                .Item("sign_sum") = fatherRow.Item("sign_sum")
                .Item("bank") = fatherRow.Item("bank")
                .Item("bank_branch") = fatherRow.Item("bank_branch")
                .Item("loanContract_num") = fatherRow.Item("loanContract_num")
                .Item("assureContract_num") = fatherRow.Item("assureContract_num")
            End With
        End If
        objProjectSignature.UpdateProjectSignature(dsTempSon)
    End Function

    '���Ʒſ��ִ����
    Public Function CopyReturnReceipt(ByVal fatherProject As String, ByVal sonProject As String)
        Dim dsTempFather, dsTempSon As DataSet
        Dim strSql As String
        Dim objLoanNotice As New LoanNotice(conn, ts)
        strSql = "{project_code='" & fatherProject & "'}"
        dsTempFather = objLoanNotice.GetLoanNoticeInfo(strSql)
        strSql = "{project_code='" & sonProject & "'}"
        dsTempSon = objLoanNotice.GetLoanNoticeInfo(strSql)
        Dim fatherRow As DataRow = dsTempFather.Tables(0).Rows(0)
        If dsTempSon.Tables(0).Rows.Count = 0 Then
            Dim sonRow As DataRow = dsTempSon.Tables(0).NewRow()
            With sonRow
                .Item("project_code") = sonProject
                .Item("bank") = fatherRow.Item("bank")
                .Item("branch_bank") = fatherRow.Item("branch_bank")
                .Item("sum") = fatherRow.Item("sum")
                .Item("term") = fatherRow.Item("term")
                .Item("start_date") = fatherRow.Item("start_date")
                .Item("end_date") = fatherRow.Item("end_date")
                .Item("return_receipt_date") = fatherRow.Item("return_receipt_date")
                .Item("register_person") = fatherRow.Item("register_person")
                .Item("register_date") = fatherRow.Item("register_date")
                .Item("create_person") = fatherRow.Item("create_person")
                .Item("create_date") = fatherRow.Item("create_date")
                .Item("affirm_person") = fatherRow.Item("affirm_person")
                .Item("affirm_date") = fatherRow.Item("affirm_date")
            End With
            dsTempSon.Tables(0).Rows.Add(sonRow)
        Else
            Dim sonRow As DataRow = dsTempSon.Tables(0).Rows(0)
            With sonRow
                .Item("project_code") = sonProject
                .Item("bank") = fatherRow.Item("bank")
                .Item("branch_bank") = fatherRow.Item("branch_bank")
                .Item("sum") = fatherRow.Item("sum")
                .Item("term") = fatherRow.Item("term")
                .Item("start_date") = fatherRow.Item("start_date")
                .Item("end_date") = fatherRow.Item("end_date")
                .Item("return_receipt_date") = fatherRow.Item("return_receipt_date")
                .Item("register_person") = fatherRow.Item("register_person")
                .Item("register_date") = fatherRow.Item("register_date")
                .Item("create_person") = fatherRow.Item("create_person")
                .Item("create_date") = fatherRow.Item("create_date")
                .Item("affirm_person") = fatherRow.Item("affirm_person")
                .Item("affirm_date") = fatherRow.Item("affirm_date")
            End With
        End If
        objLoanNotice.UpdateLoanNotice(dsTempSon)
    End Function

    '���ƽⱣ����
    Public Function CopyRefundCertificate(ByVal fatherProject As String, ByVal sonProject As String)
        Dim dsTempFather, dsTempSon As DataSet
        Dim strSql As String
        Dim objRefundCertificate As New RefundCertificate(conn, ts)
        strSql = "{project_code='" & fatherProject & "'}"
        dsTempFather = objRefundCertificate.GetRefundCertificateInfo(strSql)
        strSql = "{project_code='" & sonProject & "'}"
        dsTempSon = objRefundCertificate.GetRefundCertificateInfo(strSql)
        Dim fatherRow As DataRow = dsTempFather.Tables(0).Rows(0)
        If dsTempSon.Tables(0).Rows.Count = 0 Then
            Dim sonRow As DataRow = dsTempSon.Tables(0).NewRow()
            With sonRow
                .Item("project_code") = sonProject
                .Item("certificate_id") = fatherRow.Item("certificate_id")
                .Item("bank") = fatherRow.Item("bank")
                .Item("branch_bank") = fatherRow.Item("branch_bank")
                .Item("refund_date") = fatherRow.Item("refund_date")
                .Item("sum") = fatherRow.Item("sum")
                .Item("affirm_person") = fatherRow.Item("affirm_person")
                .Item("affirm_date") = fatherRow.Item("affirm_date")
                .Item("create_person") = fatherRow.Item("create_person")
                .Item("create_date") = fatherRow.Item("create_date")
                .Item("loanContract_num") = fatherRow.Item("loanContract_num")
                .Item("assureContract_num") = fatherRow.Item("assureContract_num")
            End With
            dsTempSon.Tables(0).Rows.Add(sonRow)
        Else
            Dim sonRow As DataRow = dsTempSon.Tables(0).Rows(0)
            With sonRow
                .Item("project_code") = sonProject
                .Item("certificate_id") = fatherRow.Item("certificate_id")
                .Item("bank") = fatherRow.Item("bank")
                .Item("branch_bank") = fatherRow.Item("branch_bank")
                .Item("refund_date") = fatherRow.Item("refund_date")
                .Item("sum") = fatherRow.Item("sum")
                .Item("affirm_person") = fatherRow.Item("affirm_person")
                .Item("affirm_date") = fatherRow.Item("affirm_date")
                .Item("create_person") = fatherRow.Item("create_person")
                .Item("create_date") = fatherRow.Item("create_date")
                .Item("loanContract_num") = fatherRow.Item("loanContract_num")
                .Item("assureContract_num") = fatherRow.Item("assureContract_num")
            End With
        End If
        objRefundCertificate.UpdateRefundCertificate(dsTempSon)
    End Function

    '��������
    Public Function finishedTask(ByVal workFlowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal finishedFlag As String, ByVal userID As String)
        'Dim strSql As String
        'Dim strMatchProjectCode As String

        'strSql = "select isnull(match_project_code,'') as match_project_code from project where project_code='" + projectID + "'"
        'Dim objCommonQuery As New CommonQuery(conn, ts)
        'Dim dsTemp As DataSet = objCommonQuery.GetCommonQueryInfo(strSql)
        'If Trim(dsTemp.Tables(0).Rows(0).Item("match_project_code")) <> "" And finishedTaskID <> "RecordReviewConclusion" Then
        '    strMatchProjectCode = Trim(dsTemp.Tables(0).Rows(0).Item("match_project_code"))

        '    '��Ӧ�Ĵ������Ŀ��ǩ���ſ�֪ͨ��
        '    If (finishedTaskID = "ValidateLoanSmall" And workFlowID = "32") Or (finishedTaskID = "ValidateLoanSmall" And workFlowID = "33") Then
        '        strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='ValidateLoan'"
        '        dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp.Tables(0).Rows.Count = 0 Then
        '            Dim wfErr As New WorkFlowErr
        '            wfErr.ThrowMustValidateLoan()
        '            Throw wfErr
        '            Exit Function
        '        End If
        '    End If

        '    '��Ӧ�Ĵ������Ŀ����ȡ����
        '    If (finishedTaskID = "LoanApplication" And workFlowID = "32") Or (finishedTaskID = "LoanApplication" And workFlowID = "33") Then
        '        strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='GuaranteeCharge'"
        '        dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp.Tables(0).Rows.Count = 0 Then
        '            Dim wfErr As New WorkFlowErr
        '            wfErr.ThrowMustGuaranteeFee()
        '            Throw wfErr
        '            Exit Function
        '        End If
        '    End If

        '    ''��Ӧ�Ĵ������Ŀ����ȡ����
        '    'If (finishedTaskID = "ServiceFeeCharge" And workFlowID = "32") Or (finishedTaskID = "ServiceFeeCharge" And workFlowID = "33") Then
        '    '    strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='GuaranteeCharge'"
        '    '    dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '    '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '    '        Dim wfErr As New WorkFlowErr
        '    '        wfErr.ThrowMustGuaranteeFee()
        '    '        Throw wfErr
        '    '        Exit Function
        '    '    End If
        '    'End If

        '    '��Ӧ��С����Ŀ��ǼǷſ��ִ
        '    If finishedTaskID = "RecordReturnReceipt" Then

        '        If workFlowID = "02" Then


        '            strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='ValidateLoanSmall'"
        '            dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)

        '            '��Ӧ��С����Ŀ��ǩ���ſ�֪ͨ��
        '            If dsTemp.Tables(0).Rows.Count = 0 Then
        '                Dim wfErr As New WorkFlowErr
        '                wfErr.ThrowMustValidateLoanSmall()
        '                Throw wfErr
        '                Exit Function
        '            End If

        '            strSql = "select workflow_id from project_task_attendee where project_code='" + strMatchProjectCode + "' and task_id='ValidateLoanSmall'"
        '            dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)

        '            'С����Ȳ���ǼǷſ��ִ�����Բ����
        '            Dim strMatchWorkflowID As String
        '            If dsTemp.Tables(0).Rows.Count <> 0 Then
        '                strMatchWorkflowID = dsTemp.Tables(0).Rows(0).Item("workflow_id")
        '            End If

        '            If strMatchWorkflowID <> "33" Then

        '                strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='RecordReturnReceipt'"
        '                dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '                If dsTemp.Tables(0).Rows.Count = 0 Then
        '                    Dim wfErr As New WorkFlowErr
        '                    wfErr.ThrowMustRecordReturnReceipt()
        '                    Throw wfErr
        '                    Exit Function
        '                End If

        '            End If

        '        End If

        '        '���Ʒſ���Ϣ
        '        CopyReturnReceipt(projectID, strMatchProjectCode)

        '    End If


        '    '���Ʒſ��¼
        '    If finishedTaskID = "LoanPetition" Then
        '        CopySignature(projectID, strMatchProjectCode)
        '    End If

        '    '��Ӧ�Ĵ������Ŀ��������ͬ
        '    If (finishedTaskID = "DraftOutContract" And workFlowID = "32") Or (finishedTaskID = "DraftOutContract" And workFlowID = "33") Then
        '        strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='DraftOutContract'"
        '        dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp.Tables(0).Rows.Count = 0 Then
        '            Dim wfErr As New WorkFlowErr
        '            wfErr.ThrowMustDraftOutContract()
        '            Throw wfErr
        '            Exit Function
        '        End If
        '    End If

        '    '����ǩԼ��Ϣ
        '    If finishedTaskID = "RecordSignature" Then
        '        CopySignature(projectID, strMatchProjectCode)
        '    End If

        '    '���ƽⱣ��Ϣ
        '    If finishedTaskID = "RecordRefundCertificate" Then
        '        CopyRefundCertificate(projectID, strMatchProjectCode)
        '    End If

        '    finishedTask(workFlowID, projectID, finishedTaskID, finishedFlag, userID, 0)

        '    strSql = "select task_id from project_task_attendee where project_code='" + strMatchProjectCode + "' and task_id='" & finishedTaskID & "'"
        '    dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '    If dsTemp.Tables(0).Rows.Count <> 0 Then
        '        finishedTask(workFlowID, strMatchProjectCode, finishedTaskID, finishedFlag, userID, 0)
        '    End If

        'Else
        '    'С���������������Ч����ʣ�����Լ����δ���ڲſ�����
        '    If workFlowID = "34" Then

        '        '���credit_project_codeΪ����Ϊ�������룬����������credit_project_codeΪ��ЧС����ȵı���
        '        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        '        Dim dsTempProject As DataSet = project.GetProjectInfo(strSql)
        '        If dsTempProject.Tables(0).Rows(0).Item("credit_project_code") Is DBNull.Value Then
        '            '��ȡ��ǰ��Ч��С�������Ŀ
        '            strSql = "select projectcode ,RemnantCredit from SmallCreditInfo where substring(projectcode,1,5)='" & projectID.Substring(0, 5) & "' order by applydate desc"
        '            dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '            If dsTemp.Tables(0).Rows.Count <> 0 Then
        '                If dsTemp.Tables(0).Rows(0).Item("RemnantCredit") > dsTempProject.Tables(0).Rows(0).Item("apply_sum") Then
        '                    dsTempProject.Tables(0).Rows(0).Item("credit_project_code") = dsTemp.Tables(0).Rows(0).Item("projectcode")
        '                    project.UpdateProject(dsTempProject)
        '                End If
        '            Else
        '                '����ЧС�����Ŷ�Ȼ��Ȳ���
        '                Dim wfErr As New WorkFlowErr
        '                wfErr.ThrowNoSmallCredit()
        '                Throw wfErr
        '                Exit Function
        '            End If
        '        End If
        '    End If

        '    '������ǰС����Ŀ�軹���ſɷſ�
        '    If workFlowID = "02" Then
        '        strSql = "select ServiceType,relate_project_code from queryProjectInfo where project_code='" & projectID & "'"
        '        dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp.Tables(0).Rows(0).Item("ServcieType") = "��ǰС��" Then
        '            If IsDBNull(dsTemp.Tables(0).Rows(0).Item("relate_project_code")) = False Then
        '                strSql = "select guaranting_sum from queryProjectInfo where project_code='" & dsTemp.Tables(0).Rows(0).Item("relate_project_code") & "'"
        '                dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '                If dsTemp.Tables(0).Rows(0).Item("guaranting_sum") <> 0 Then
        '                    Dim wfErr As New WorkFlowErr
        '                    wfErr.ThrowRelateProjectRefund()
        '                    Throw wfErr
        '                    Exit Function
        '                End If
        '            End If
        '        End If
        '    End If

        '    finishedTask(workFlowID, projectID, finishedTaskID, finishedFlag, userID, 0)

        'End If

        Return finishedTask(workFlowID, projectID, finishedTaskID, finishedFlag, userID, 0)

    End Function
    Public Function finishedTask(ByVal workFlowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal finishedFlag As String, ByVal userID As String, ByVal flag As Integer)
        Dim startTime As DateTime
        Dim tmpTaskName, tmpTaskType, tmpTaskPhase, tmpTaskStatus, tmpStartMode As String
        Dim i, j, k As Integer
        Dim newRow As DataRow

        '������������ύ��,

        '1�������ǰ����ID�����ڣ��׳��������񲼴����쳣
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            'Dim wfErr As New WorkFlowErr
            'wfErr.ThrowNotExistTaskErr()
            'Throw wfErr
            Exit Function
        Else


            '��ȡ��ǰ���������,�������ͺͿ�ʼʱ��
            tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
            tmpTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))
            startTime = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_time")), Now, dsTempTask.Tables(0).Rows(0).Item("start_time"))
            tmpTaskPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))
            tmpTaskStatus = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_status")), "", dsTempTask.Tables(0).Rows(0).Item("project_status")))
            tmpStartMode = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_mode")), "", dsTempTask.Tables(0).Rows(0).Item("start_mode")))

            '��ȡ��Ŀ�׶κ���Ŀ״̬
            Dim tmpProjectPhase, tmpProjectStatus As String
            Dim dsTempProject As DataSet
            'strSql = "{project_code=" & "'" & projectID & "'" & "}"
            'dsTempProject = project.GetProjectInfo(strSql)

            ''�쳣����  
            'If dsTempProject.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            '    Throw wfErr
            'End If

            'If tmpTaskPhase <> "" Then
            '    dsTempProject.Tables(0).Rows(0).Item("phase") = tmpTaskPhase
            'End If

            'project.UpdateProject(dsTempProject)

            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)

            tmpProjectPhase = Trim(IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("phase")), "", dsTempProject.Tables(0).Rows(0).Item("phase")))
            tmpProjectStatus = Trim(IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("status")), "", dsTempProject.Tables(0).Rows(0).Item("status")))

            '2�������ǰ����ID״̬Ϊ��W�����׳������ύ��ͣ�������쳣
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
                Throw wfErr
            End If

            Dim isWaiting As String = Trim(IIf(IsDBNull(dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status")))
            If isWaiting = "W" Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowWaitingTaskErr()
                Throw wfErr
                Exit Function
            End If

            '3�������������=��ISLAND����ȡϵͳ���ʱ����Ϊ��ǰ��������ʱ�䣬���أ�
            If tmpTaskType = "ISLAND" Then

                '��ȡ������־�и���Ŀ��������AUTO����Ϊ0����־
                Dim Worklog As New WorkLog(conn, ts)
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and auto=0}"
                Dim dsWorklog As DataSet = Worklog.GetWorkLogInfo(strSql)

                For i = 0 To dsWorklog.Tables(0).Rows.Count - 1

                    newRow = dsWorklog.Tables(0).Rows(i)
                    With newRow

                        '������״̬��ΪF,AUTO��Ϊ1,��Ŀ�׶κ�״̬��Ϊ��Ӧֵ
                        .Item("task_status") = "F"
                        .Item("project_phase") = tmpProjectPhase

                        '������������״̬����Ѹ�״̬��¼��������־�У��������Ŀ��״̬��¼��������־��
                        If tmpTaskStatus = "" Then

                            .Item("project_status") = tmpProjectStatus

                            ''��ӹ�����־
                            'AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpProjectStatus, tmpStartMode)

                        Else
                            .Item("project_status") = tmpTaskStatus

                            'AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpTaskStatus, tmpStartMode)
                        End If

                    End With
                Next

                Worklog.UpdateWorkLog(dsWorklog)

                ''���TaskID����������߾���ɴ�����()
                'Dim isDone As Boolean = True
                'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                '    If Trim(IIf(IsDBNull(dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status"))) <> "F" Then
                '        isDone = False
                '    End If
                'Next

                ''   ������Project_Code= ProjectID ��StartupTask= TaskID��Status=��P��������Project_Track�����Status��Ϊ��F����
                'If isDone Then
                '    strSql = "{workflow_id=" & "'" & workFlowID & "'" & " and project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & finishedTaskID & "'" & " and isnull(status,'')='P'}"
                '    Dim dsTempProjectTrack As DataSet = WfProjectTrack.GetWfProjectTrackInfo(strSql)
                '    For i = 0 To dsTempProjectTrack.Tables(0).Rows.Count - 1
                '        dsTempProjectTrack.Tables(0).Rows(i).Item("Status") = "F"
                '    Next
                '    WfProjectTrack.UpdateWfProjectTrack(dsTempProjectTrack)
                'End If




                Exit Function
            End If


            '��ȡϵͳʱ��
            Dim sysTime As DateTime = Now

            '2008-8-20 yjf add ���ӹ���������:ͬһ��������ж��˲���,��ֻҪ����һ���ύ���������
            If tmpTaskType = "OPT" Then

                '7��[��Ч�����������]���������ǰ����ProjectID��TaskID��EmployeeID��״̬��Ϊ��F�����
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'}"
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '�쳣����  
                If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
                    Throw wfErr
                End If

                Dim tempOptPerson As String
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    tempOptPerson = dsTempTaskAttendee.Tables(0).Rows(i).Item("attend_person")
                    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
                    '����AACKMassage��ProjectID��TaskID��EmployeeID���Զ�ȷ��������Ϣ��
                    AACKMassage(workFlowID, projectID, finishedTaskID, tempOptPerson)
                    '��Ա���������ʱ����Ϊϵͳʱ��
                    dsTempTaskAttendee.Tables(0).Rows(0).Item("end_time") = sysTime
                Next


            Else
                '7��[��Ч�����������]���������ǰ����ProjectID��TaskID��EmployeeID��״̬��Ϊ��F�����
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and attend_person=" & "'" & userID & "'" & "}"
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '�쳣����  
                If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
                    Throw wfErr
                End If

                dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status") = "F"

                '8,	����AACKMassage��ProjectID��TaskID��EmployeeID���Զ�ȷ��������Ϣ��
                AACKMassage(workFlowID, projectID, finishedTaskID, userID)

                '��Ա���������ʱ����Ϊϵͳʱ��
                dsTempTaskAttendee.Tables(0).Rows(0).Item("end_time") = sysTime

            End If

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee) '���²���

            '9����ӹ�����־(���������Ͳ�Ϊ������,���״̬Ϊ�յĹ�����־,��������ת�����״̬)
            'If isTaskStatusEqualP(projectID, finishedTaskID) Then
            If tmpTaskType <> "END" Then
                AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, "", tmpStartMode)
            End If
            'End If

            '10�������ProjectID��TaskID������������״̬��Ϊ��F�� ����ʱ�����еĵ�ǰ����ID��״̬��Ϊ��E����
            '���򣬷��أ�

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            If flag = 0 Then
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '�ж��Ƿ���������״̬��Ϊ��F��
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    If dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") <> "F" Then
                        Exit Function
                    End If
                Next
            End If
            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            ''�ж��Ƿ���������״̬��Ϊ��F��
            'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            '    If dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") <> "F" Then
            '        Exit Function
            '    End If
            'Next

            '����ʱ�����еĵ�ǰ����ID��״̬��Ϊ��E��
            Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
            If dsTempTimingTask.Tables(0).Rows.Count <> 0 Then
                For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                    dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                Next
            End If

            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

            '[�ύ�ֶ���������]�����ǰ�����start_mode=��manual������start_mode�ÿգ�����
            If tmpStartMode = "manual" Then
                dsTempTask.Tables(0).Rows(0).Item("start_mode") = ""
                WfProjectTask.UpdateWfProjectTask(dsTempTask)
                Exit Function
            End If

            '2005-04-27 yjf �޸ģ��ֶ������ύ���ı���Ŀ״̬
            '����Ŀ�Ľ׶���Ϊ�������Ӧ�׶�
            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)

            '�쳣����  
            If dsTempProject.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                Throw wfErr
            End If

            If tmpTaskPhase <> "" Then
                dsTempProject.Tables(0).Rows(0).Item("phase") = tmpTaskPhase
            End If

            project.UpdateProject(dsTempProject)

            '11�������ǰ�����ṩ���̹��ߣ��������̹���
            Dim tmpFlowTools As String
            Dim args As Object() = {conn, ts}
            If Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("flow_tool")), "", dsTempTask.Tables(0).Rows(0).Item("flow_tool"))) <> "" Then
                tmpFlowTools = dsTempTask.Tables(0).Rows(0).Item("flow_tool")
                tmpFlowTools = "BusinessRules." & tmpFlowTools
                tmpFlowTools = tmpFlowTools.Trim

                '��̬�����ӿڶ���
                Dim t As System.Type = System.Type.GetType(tmpFlowTools)
                Dim iFlowTools As IFlowTools = Activator.CreateInstance(t, args)
                iFlowTools.UseFlowTools(workFlowID, projectID, finishedTaskID, finishedFlag, userID)

            End If

            '4����ȡ��ǰ�����ת�������ת��������¼��
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            Dim dsTempTaskTransfer As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '5�� ��ȡת������Ϊ���ת������
            Dim nextTaskID, tmpTransCondition, tmpTransPhase, tmpTransStatus As String
            Dim dsConditionTrue As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")
            For i = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
                nextTaskID = Trim(dsTempTaskTransfer.Tables(0).Rows(i).Item("next_task"))
                tmpTransCondition = Trim(dsTempTaskTransfer.Tables(0).Rows(i).Item("transfer_condition"))
                tmpTransStatus = Trim(IIf(IsDBNull(dsTempTaskTransfer.Tables(0).Rows(i).Item("project_status")), "", dsTempTaskTransfer.Tables(0).Rows(i).Item("project_status")))

                '�ж������Ƿ�Ϊ��
                If CompareExpression(workFlowID, projectID, finishedTaskID, finishedFlag, tmpTransCondition) Then

                    '����ת������Ϊ���ת������
                    newRow = dsConditionTrue.Tables(0).NewRow()
                    With newRow
                        .Item("workflow_id") = ""
                        .Item("project_code") = projectID
                        .Item("project_status") = tmpTransStatus
                        .Item("task_id") = finishedTaskID
                        .Item("next_task") = nextTaskID
                        .Item("transfer_condition") = tmpTransCondition
                        .Item("project_status") = tmpTransStatus
                    End With
                    dsConditionTrue.Tables(0).Rows.Add(newRow)
                End If
            Next

            '6�����ת������Ϊ���ת������Ϊ��
            If dsConditionTrue.Tables(0).Rows.Count = 0 Then

                '�����ǰ����Ͳ��ǽ��������׳��ύ����Ľ����Ч����,���ء�
                If tmpTaskType <> "END" Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowInvalidSubmit()
                    Throw wfErr
                    Exit Function
                End If
            End If



            '12�������ǰ���������Ϊ�������������أ�
            If tmpTaskType = "END" Then

                '������������״̬����Ѹ�״̬��¼��������־�У��������Ŀ��״̬��¼��������־��
                ' If isTaskStatusEqualP(projectID, finishedTaskID) Then
                If tmpTaskStatus = "" Then
                    '��ӹ�����־
                    AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpProjectStatus, tmpStartMode)
                Else
                    AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpTaskStatus, tmpStartMode)
                End If
                'End If

                Exit Function

            End If

            '13��[���������һ���к�̻������]��ÿ��ת������Ϊ���ת�����񣨶������

            '   ����㼯����Ϊ��AND������ȡת�������ǰ���״̬��

            Dim tmpNextTaskID As String
            Dim tmpPreTaskID As String
            Dim dsTempPreTaskStatus, dsTempWorkLog As DataSet
            Dim tmpApplyTool As String

            '��FiliterJeeDummyTask��ProjectID��ShiftTaskSet����ȡת������Ϊ���ҵ������ʵ����
            dsConditionTrue = FiliterJeeDummyTask(workFlowID, projectID, dsConditionTrue, finishedFlag, userID)

            '���ʵ���񼯼�¼Ϊ��,֤��Ϊ��Ŀ����,��������־�и�����״̬��Ϊ��Ŀ״̬
            If dsConditionTrue.Tables(0).Rows.Count = 0 Then

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and project_status=''}"
                dsTempWorkLog = Worklog.GetWorkLogInfo(strSql)

                '���»�ȡ��Ŀ״̬
                strSql = "{project_code=" & "'" & projectID & "'" & "}"
                dsTempProject = project.GetProjectInfo(strSql)
                tmpProjectStatus = Trim(IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("status")), "", dsTempProject.Tables(0).Rows(0).Item("status")))

                For j = 0 To dsTempWorkLog.Tables(0).Rows.Count - 1
                    dsTempWorkLog.Tables(0).Rows(j).Item("project_status") = tmpProjectStatus
                Next
                Worklog.UpdateWorkLog(dsTempWorkLog)

            End If

            '��������ת������Ϊ���ת������
            Dim tmpNextTaskType, mergeRelation As String
            Dim dsPreTaskMode As DataSet
            Dim tmpPreTaskMode As String

            '2005-09-13 yjf add �޸����ж����������ʱ,���ڵ�һ����������ǰ������δ��ɶ����µڶ����������񲻴�������
            Dim isFinishedPreTask As Boolean

            For i = 0 To dsConditionTrue.Tables(0).Rows.Count - 1

                isFinishedPreTask = True

                tmpNextTaskID = dsConditionTrue.Tables(0).Rows(i).Item("next_task")
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpNextTaskID & "'" & "}"
                dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                '�쳣����  
                If dsTempTask.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                    Throw wfErr
                End If


                tmpNextTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))
                mergeRelation = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("merge_relation")), "", dsTempTask.Tables(0).Rows(0).Item("merge_relation"))

                tmpTransStatus = Trim(IIf(IsDBNull(dsConditionTrue.Tables(0).Rows(i).Item("project_status")), "", dsConditionTrue.Tables(0).Rows(i).Item("project_status")))

                '����AddTaskTrackRecord��ProjectID,Workflow_id,TaskID,StartupTask����¼����������Ϣ;
                AddTaskTrackRecord("", projectID, finishedTaskID, tmpNextTaskID)

                '���ת���������Ŀ״̬����ֵ�ǿ�
                '   ����Ŀ��״̬�޸�Ϊת���������Ŀ״̬����ֵ��
                strSql = "{project_code=" & "'" & projectID & "'" & "}"
                dsTempProject = project.GetProjectInfo(strSql)

                '�쳣����  
                If dsTempProject.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                    Throw wfErr
                End If

                If tmpTransStatus <> "" Then
                    dsTempProject.Tables(0).Rows(0).Item("status") = tmpTransStatus
                    project.UpdateProject(dsTempProject)
                End If

                '��ȡ��Ŀ��״̬��
                '�ڹ�����־��TaskID�������Ŀ״̬Ϊ�յļ�¼����Ŀ״̬��Ϊ��Ŀ��״ֵ̬;
                dsTempProject = project.GetProjectInfo(strSql)

                '�쳣����  
                If dsTempProject.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                    Throw wfErr
                End If

                tmpProjectStatus = Trim(IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("status")), "", dsTempProject.Tables(0).Rows(0).Item("status")))
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and project_status=''}"
                dsTempWorkLog = Worklog.GetWorkLogInfo(strSql)

                For j = 0 To dsTempWorkLog.Tables(0).Rows.Count - 1
                    dsTempWorkLog.Tables(0).Rows(j).Item("project_status") = tmpProjectStatus
                Next
                Worklog.UpdateWorkLog(dsTempWorkLog)

                '��ȡÿһ��ת������Ϊ���ת������ID

                If mergeRelation = "AND" Then
                    '����㼯����Ϊ��AND������ȡת�������ǰ���״̬��

                    strSql = "{project_code=" & "'" & projectID & "'" & " and next_task=" & "'" & tmpNextTaskID & "'" & "}"
                    '��ȡת�������ǰ�����
                    dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    '����ת�������ǰ�����
                    For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1

                        tmpPreTaskID = dsTempTaskTransfer.Tables(0).Rows(j).Item("task_id")
                        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpPreTaskID & "'" & "}"

                        '��ȡÿ��ǰ��������ɫ�������״̬��
                        dsTempPreTaskStatus = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                        '����ǰ��������״̬���Ƿ���ΪP��������ģʽ��ΪManual��
                        For k = 0 To dsTempPreTaskStatus.Tables(0).Rows.Count - 1

                            '��ȡǰ�����������ģʽ
                            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpPreTaskID & "'" & "}"
                            dsPreTaskMode = WfProjectTask.GetWfProjectTaskInfo(strSql)

                            '�쳣����  
                            If dsPreTaskMode.Tables(0).Rows.Count = 0 Then
                                Dim wfErr As New WorkFlowErr
                                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                                Throw wfErr
                            End If

                            tmpPreTaskMode = IIf(IsDBNull(dsPreTaskMode.Tables(0).Rows(0).Item("start_mode")), "", dsPreTaskMode.Tables(0).Rows(0).Item("start_mode"))

                            If IIf(IsDBNull(dsTempPreTaskStatus.Tables(0).Rows(k).Item("task_status")), "", dsTempPreTaskStatus.Tables(0).Rows(k).Item("task_status")) <> "F" And tmpPreTaskMode <> "manual" Then
                                '2005-09-13 yjf add �޸����ж����������ʱ,���ڵ�һ����������ǰ������δ��ɶ����µڶ����������񲻴�������
                                isFinishedPreTask = False
                                'Exit Function
                            End If
                        Next
                    Next

                    '2005-09-13 yjf add �޸����ж����������ʱ,���ڵ�һ����������ǰ������δ��ɶ����µڶ����������񲻴�������
                    If isFinishedPreTask Then

                        '����ת������
                        StartupTask(workFlowID, projectID, tmpNextTaskID, "", "", finishedTaskID, userID)

                    End If

                Else
                    If mergeRelation = "XOR" Then
                        '����ת������󷵻�
                        StartupTask(workFlowID, projectID, tmpNextTaskID, "", "", finishedTaskID, userID)
                        Exit Function
                    Else
                        '����ת������
                        StartupTask(workFlowID, projectID, tmpNextTaskID, "", "", finishedTaskID, userID)
                    End If
                End If

                'If tmpNextTaskType = "AUTO" Then

                '    '��ȡӦ�ù���
                '    strSql = "{project_code=" & "'" & projectID & "'"   & " and task_id=" & "'" & tmpNextTaskID & "'" & "}"
                '    dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
                '    tmpApplyTool = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("apply_tool")), "", dsTempTask.Tables(0).Rows(0).Item("apply_tool")))

                '    If tmpApplyTool <> "" Then
                '        '���ù���
                '        tmpApplyTool = "BusinessRules." & tmpApplyTool
                '        Dim t As System.Type = System.Type.GetType(tmpApplyTool)
                '        Dim iApplyTools As IApplyTools = Activator.CreateInstance(t, args)
                '        iApplyTools.UseApplyTools()

                '    End If

                '    '������������ִ����ɹ��ߺ󣬵��ô˷���������ת
                '    VTask(workFlowID, projectID, tmpNextTaskID, userID)

                'End If
            Next

        End If

    End Function

    '�жϹ������Ƿ����
    Public Function isSuspendProcess(ByVal projectID As String) As Boolean
        Dim strSql As String
        Dim dsTempAttendee As DataSet
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='W'" & "}"
        dsTempAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If dsTempAttendee.Tables(0).Rows.Count <> 0 Then
            isSuspendProcess = True
        Else
            isSuspendProcess = False
        End If
    End Function

    '���̹���
    Public Function suspendProcess(ByVal projectID As String, ByVal delayDay As Integer)

        '��	��ָ������������״̬Ϊ��P��������״̬��Ϊ��W����
        Dim i, j As Integer
        Dim sysTime As DateTime = Now
        Dim strSql As String
        Dim dsTempAttendee, dsTempTask, dstTempTimingTask As DataSet
        Dim tmpTaskID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='P'" & "}"
        dsTempAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttendee.Tables(0).Rows.Count - 1

            dsTempAttendee.Tables(0).Rows(i).Item("task_status") = "W"

            '��	�������ͣ��ʼʱ����Ϊϵͳʱ�䣻
            tmpTaskID = Trim(dsTempAttendee.Tables(0).Rows(i).Item("task_id"))
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
            For j = 0 To dsTempTask.Tables(0).Rows.Count - 1
                dsTempTask.Tables(0).Rows(j).Item("pause_start_time") = sysTime
            Next
            WfProjectTask.UpdateWfProjectTask(dsTempTask)

        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttendee)


        '��	��ȡ��ʱ�������ָ����������ʱ����Ķ�ʱ����ΪA��״̬Ϊ��P���Ķ�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and type='A' and status='P'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        '��	����ʱ����״̬��Ϊ��W��
        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "W"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)

        '����ʱ������еĻָ������������״̬��Ϊ'P',��ʼʱ����Ϊ��ǰʱ��+���ʱ��
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='WakeProject'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
            dstTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Day, delayDay, Now)
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)

        '��Ŀ��ͣ:���ϼ����ܷ�����Ϣ
        sendMessageToManager(projectID, delayDay)


    End Function

    '���ָ̻�
    Public Function resumeProcess(ByVal projectID As String)
        '��	��������ȡ��Ŀ����ָ��������״̬Ϊ��W��������
        Dim i, j As Integer
        Dim sysTime As DateTime = Now
        Dim strSql As String
        Dim dsTempAttendee, dsTempTask, dstTempTimingTask As DataSet
        Dim tmpTaskID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='W'" & "}"
        dsTempAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '��	��״̬Ϊ��W��������״̬��Ϊ��P����
        For i = 0 To dsTempAttendee.Tables(0).Rows.Count - 1
            dsTempAttendee.Tables(0).Rows(i).Item("task_status") = "P"

            '��	��ȡϵͳʱ�䣻
            '��	��ͣ����ʱ����Ϊϵͳʱ�䣻
            tmpTaskID = Trim(dsTempAttendee.Tables(0).Rows(i).Item("task_id"))
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
            For j = 0 To dsTempTask.Tables(0).Rows.Count - 1
                dsTempTask.Tables(0).Rows(j).Item("pause_end_time") = sysTime
            Next
            WfProjectTask.UpdateWfProjectTask(dsTempTask)
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttendee)

        '��	�ڶ�ʱ������н�ָ����������ʱ����״̬Ϊ��W��������ʼʱ���Ϊ��ʼʱ��+(��ͣ����ʱ��-��ͣ��ʼʱ��)�� 
        strSql = "{project_code=" & "'" & projectID & "'" & " and status='W'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        Dim tmpWorkFlowID As String
        Dim tmpStartTime, tmpPauseEndTime, tmpPauseStartTime As DateTime

        '��	��ָ����������ʱ����״̬Ϊ��W��������״̬��Ϊ��P���� 
        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            tmpWorkFlowID = dstTempTimingTask.Tables(0).Rows(i).Item("workflow_id")
            tmpTaskID = dstTempTimingTask.Tables(0).Rows(i).Item("task_id")
            strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '�쳣����  
            If dsTempTask.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                Throw wfErr
            End If

            tmpStartTime = CDate(FormatDateTime(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_time")), Now, dsTempTask.Tables(0).Rows(0).Item("start_time")), DateFormat.ShortDate))
            tmpPauseEndTime = CDate(FormatDateTime(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("pause_end_time")), Now, dsTempTask.Tables(0).Rows(0).Item("pause_end_time"))))
            tmpPauseStartTime = CDate(FormatDateTime(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("pause_start_time")), Now, dsTempTask.Tables(0).Rows(0).Item("pause_start_time"))))

            dstTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Day, DateDiff(DateInterval.Day, tmpPauseStartTime, tmpPauseEndTime), tmpStartTime)
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)

        '����ʱ������еĻָ������������״̬��Ϊ'E'
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='WakeProject'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)


    End Function

    '�������
    Public Function rollbackTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal userID As String, ByVal rollbackMsg As String)

        Dim i, j As Integer
        Dim strSql As String
        Dim dsTask, dsProjectTrack, dsWorkLog, dsProject, dsAttend, dsTimingTask As DataSet
        Dim tmpStartMode, tmpRollBackTask, tmpFinshedTask, tmpWorklogPhase, tmpWorklogStatus, tmpStartupTask, tmpTaskStatus As String
        Dim iRangUp As Integer

        '��	���Project_Task.Start_Mode=��manual��,��ʾ�����ܻ����ֹ����������񣡡������أ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����  
        If dsTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTask.Tables(0))
            Throw wfErr
        End If

        tmpStartMode = IIf(IsDBNull(dsTask.Tables(0).Rows(0).Item("start_mode")), "", dsTask.Tables(0).Rows(0).Item("start_mode"))
        If tmpStartMode = "manual" Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowRollBackManualTaskErr()
            Throw wfErr
        End If

        '��	��ȡWorkflow_id=������ID��StartupTask= TaskID��Status=��P����Project_Track����Serial_Num��С�ߣ���
        strSql = "{project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & taskID & "'" & " and isnull(status,'')='P'}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        '�쳣����  
        If dsProjectTrack.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectTrack.Tables(0))
            Throw wfErr
        End If

        '��	��ȡProject_Track�����FinishedTask���ԣ�RollBackTask����
        tmpRollBackTask = Trim(dsProjectTrack.Tables(0).Rows(0).Item("FinishedTask"))

        '��	��ȡProject_Track�����Serial_Num����(RangeUp);
        iRangUp = dsProjectTrack.Tables(0).Rows(0).Item("serial_num")

        '��	��ȡ���������������������λ�ã�Serial_Num��
        strSql = "{project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & tmpRollBackTask & "'" & " and serial_num<" & "'" & iRangUp & "'" & " order by serial_num desc}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        '�쳣����  
        If dsProjectTrack.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectTrack.Tables(0))
            Throw wfErr
        End If

        iRangUp = dsProjectTrack.Tables(0).Rows(0).Item("serial_num")
        tmpFinshedTask = dsProjectTrack.Tables(0).Rows(0).Item("FinishedTask")

        '�� ��Project_Track�����Status������Ϊ��P����
        dsProjectTrack.Tables(0).Rows(0).Item("Status") = "P"
        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

        '�ڹ�����־��ȡProject_Track. FinishedTask����Ŀ�׶κ���Ŀ״̬;
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpFinshedTask & "'" & " order by finish_time desc}"
        dsWorkLog = WorkLog.GetWorkLogInfo(strSql)

        '�쳣����  
        If dsWorkLog.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsWorkLog.Tables(0))
            Throw wfErr
        End If

        tmpWorklogPhase = IIf(IsDBNull(dsWorkLog.Tables(0).Rows(0).Item("project_phase")), "", dsWorkLog.Tables(0).Rows(0).Item("project_phase"))
        tmpWorklogStatus = IIf(IsDBNull(dsWorkLog.Tables(0).Rows(0).Item("project_status")), "", dsWorkLog.Tables(0).Rows(0).Item("project_status"))

        '����Ŀ״̬�ͽ׶θ���ΪProject_Track. FinishedTask����Ŀ�׶κ���Ŀ״̬
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsProject = project.GetProjectInfo(strSql)

        '�쳣����  
        If dsProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProject.Tables(0))
            Throw wfErr
        End If

        dsProject.Tables(0).Rows(0).Item("phase") = tmpWorklogPhase
        dsProject.Tables(0).Rows(0).Item("status") = tmpWorklogStatus
        project.UpdateProject(dsProject)


        '��	��RollBackTask��������������RollBackSet��
        Dim ArrRollBackTask As New ArrayList
        ArrRollBackTask.Add(tmpRollBackTask)


        '��	����SERIAL-NUM> RangeUp��ÿ��Project_Track����
        strSql = "{project_code=" & "'" & projectID & "'" & " and serial_num>" & "'" & iRangUp & "'" & " order by serial_num}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1

            tmpFinshedTask = dsProjectTrack.Tables(0).Rows(i).Item("FinishedTask")
            For j = 0 To ArrRollBackTask.Count - 1
                '���Project_Track. FinishedTask IN RollBackSet
                If tmpFinshedTask = ArrRollBackTask.Item(j) Then
                    '��Project_Track. StartupTask��������������RollBackSet;
                    tmpStartupTask = dsProjectTrack.Tables(0).Rows(i).Item("StartupTask")
                    ArrRollBackTask.Add(tmpStartupTask)

                    'ɾ��Project_Track����
                    dsProjectTrack.Tables(0).Rows(i).Delete()

                    Exit For

                End If
            Next

        Next

        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

        '���ڻ�������RollBackSet�е�ÿ������RollBackSet(i)
        For i = 0 To ArrRollBackTask.Count - 1

            '���RollBackSet(i) = rollbackTask
            If ArrRollBackTask(i) = tmpRollBackTask Then
                '   ����startupTask(ģ��ID����ĿID��RollBackTask)������������;
                StartupTask(workFlowID, projectID, ArrRollBackTask(i), rollbackMsg, userID)

            Else
                '����()

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & ArrRollBackTask(i) & "'" & "}"
                dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                '      ������״̬��Ϊ����;

                For j = 0 To dsAttend.Tables(0).Rows.Count - 1
                    dsAttend.Tables(0).Rows(j).Item("task_status") = ""
                Next
                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

                dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

                '      ������Ķ�ʱ����״̬��Ϊ��E��;
                For j = 0 To dsTimingTask.Tables(0).Rows.Count - 1
                    dsTimingTask.Tables(0).Rows(j).Item("status") = "E"
                Next
                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

            End If
        Next

    End Function

    Private Function GetRollbackTask(ByVal dsPreTask As DataSet, ByVal dsFinishedTask As DataSet) As String
        Dim i, j As Integer
        Dim tmpFinishedTask, tmpRollBackTask As String
        For i = 0 To dsFinishedTask.Tables(0).Rows.Count - 1
            tmpFinishedTask = Trim(dsFinishedTask.Tables(0).Rows(i).Item("task_id"))
            For j = 0 To dsPreTask.Tables(0).Rows.Count - 1
                tmpRollBackTask = Trim(dsPreTask.Tables(0).Rows(j).Item("task_id"))
                If tmpFinishedTask = tmpRollBackTask Then
                    Return tmpRollBackTask
                End If
            Next
        Next
    End Function

    '��ֹ����
    Public Function cancelProcess(ByVal projectID As String)

        '��	ɾ����������˱�����Ŀ���������
        Dim i As Integer
        Dim dsTempTask, dsTempAttend, dsTempTimingTask, dsTempTrans, dsTempProject, dsWorkLog, dsProjectTrack As DataSet
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & "}"
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

        '��	ɾ����ʱ������ָ����Ŀ����Ķ�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


        '��	ɾ��������־�е�����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsWorkLog = WorkLog.GetWorkLogInfo(strSql)
        For i = 0 To dsWorkLog.Tables(0).Rows.Count - 1
            dsWorkLog.Tables(0).Rows(i).Delete()
        Next
        WorkLog.UpdateWorkLog(dsWorkLog)


        'ɾ��������ٱ��еļ�¼
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)
        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1
            dsProjectTrack.Tables(0).Rows(i).Delete()
        Next
        WfProjectTrack.UpdateWfProjectTrack(dsTempTimingTask)

        '�� ɾ��������е�����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
        For i = 0 To dsTempTask.Tables(0).Rows.Count - 1

            dsTempTask.Tables(0).Rows(i).Delete()
        Next

        WfProjectTask.UpdateWfProjectTask(dsTempTask)


        '������Ŀ������ʶisliving=0
        dsTempProject = project.GetProjectInfo(strSql)

        '�쳣����  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        dsTempProject.Tables(0).Rows(0).Item("isliving") = 0

        '������Ŀ��״̬Ϊ��Ŀ�׶�+"�ݻ�"
        dsTempProject.Tables(0).Rows(0).Item("status") = dsTempProject.Tables(0).Rows(0).Item("phase") & "�ݻ�"
        project.UpdateProject(dsTempProject)


    End Function


    '�ֹ���������WorkflowID��ProjectID��TaskID��
    Public Function StartTaskByManual(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String)
        Dim strSql As String
        Dim dsTempTask As DataSet
        '��	��������ȡ����ָ�������񣬽�start_mode��Ϊ��manual����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����  
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        dsTempTask.Tables(0).Rows(0).Item("start_mode") = "manual"
        WfProjectTask.UpdateWfProjectTask(dsTempTask)
        '��	����StartupManualTask(WorkflowID��ProjectID��TaskID������)��
        StartupManualTask(workflowID, projectID, taskID, "", "")
    End Function

    '��ȡ��������WorkflowID��ProjectID��
    Public Function GetAllBusinessTasks(ByVal workflowID As String, ByVal projectID As String) As DataSet
        Dim strSql As String
        Dim dsTempAttend As DataSet
        '��	�������ɫ���ȡ�����WorkflowID��ProjectIDƥ��������������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        '��	�������������б�
        Return dsTempAttend
    End Function


    'ɾ��������
    Public Function deleteProcess(ByVal projectID As String)

    End Function

    '��������
    Public Function modifiyProcess(ByVal projectID As String)

    End Function

    '��ѯ��Ϣ��Ϣ
    Public Function LookUpMessage(ByVal strCondition_ProjectMessage As String) As DataSet
        'Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and accepter=" & "'" & userID & "'" & "}"
        Dim dsTempProjectMessage As DataSet = WfProjectMessages.GetWfProjectMessagesInfo(strCondition_ProjectMessage)
        Return dsTempProjectMessage
    End Function

    '��ѯ�����е�����
    Public Function LookUpWorking(ByVal projectID As String, ByVal userID As String)

    End Function

    '��ѯ�����е�����
    Public Function LookUpWorking(ByVal userID As String) As DataSet

        ''����״̬Ϊ��P�������е���Ŀ��ȡ������ĿID������ID
        ''Dim strSql As String = "{attend_person=" & "'" & userID & "'" & " and task_status='P' order by task_id}"
        'Dim strSql As String = " SELECT dbo.project_task.*, dbo.project_task_attendee.role_id " & _
        '                       " FROM dbo.project_task INNER JOIN " & _
        '                       " dbo.project_task_attendee ON " & _
        '                       " dbo.project_task.project_code = dbo.project_task_attendee.project_code AND " & _
        '                       " dbo.project_task.workflow_id = dbo.project_task_attendee.workflow_id AND " & _
        '                       " dbo.project_task.task_id = dbo.project_task_attendee.task_id AND " & _
        '                       " dbo.project_task_attendee.attend_person=" & "'" & userID & "'" & " and dbo.project_task_attendee.task_status='P' order by dbo.project_task.task_id "
        ''Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        ''Dim dsTemp As DataSet = Me.commQuery.GetCommonQueryInfo(strSql)
        ''Dim dsTask As DataSet = WfProjectTask.GetWfProjectTaskInfo("null")
        ''Dim projectID, taskID As String

        ''Dim i As Integer

        '''����ָ����Ŀ������������б�
        ''For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        ''    projectID = Trim(dsTemp.Tables(0).Rows(i).Item("project_code"))
        ''    taskID = dsTemp.Tables(0).Rows(i).Item("task_id")
        ''    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id =" & "'" & taskID & "'" & " }"
        ''    dsTask.Merge(WfProjectTask.GetWfProjectTaskInfo(strSql))
        ''Next

        'Dim dsTask As DataSet = Me.commQuery.GetCommonQueryInfo(strSql)

        ''���������б�
        'Return dsTask

        Dim dsTask As DataSet = Me.commQuery.LookUpWorking(userID)
        Return dsTask

    End Function

    '��ѯ�����е�����
    Public Function LookUpWorkingEx(ByVal sql_Condition As String) As DataSet

        '����״̬Ϊ��P�������е���Ŀ��ȡ������ĿID������ID
        Dim strSql As String = sql_Condition
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim dsTask As DataSet = WfProjectTask.GetWfProjectTaskInfo("null")
        Dim projectID, taskID As String

        Dim i As Integer

        '����ָ����Ŀ������������б�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            projectID = Trim(dsTemp.Tables(0).Rows(i).Item("project_code"))
            taskID = dsTemp.Tables(0).Rows(i).Item("task_id")
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id =" & "'" & taskID & "'" & "}"
            dsTask.Merge(WfProjectTask.GetWfProjectTaskInfo(strSql))
        Next

        '���������б�
        Return dsTask

    End Function

    '��ѯ����״̬
    Public Function LookUpStatus(ByVal projectID As String)

    End Function

    '�Ƚ��жϱ��ʽ�Ƿ�Ϊ��
    Private Function CompareExpression(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean
        '������������ӿ�
        Dim iCondition As ICondition

        'Select Case taskID

        '    Case "ReviewFeeCharge", "CashlossReview"   '����������Ƿ����֧��
        '        iCondition = New ImplIncomePayout(conn, ts)
        '        'Case "CheckApplyTimes" '�������>3
        '        '    iCondition = New ImplCommon()
        '        '    Return iCondition.GetResult(workFlowID, projectID, taskID, ".T.", transCondition)
        '        'Case "ValidateReviewConclusion" '������Ƿ���
        '        '    iCondition = New ImplTrialFee(conn, ts)
        '        ''Case "GuaranteeCharge" '�����������Ƿ����֧��
        '        ''    iCondition = New ImplGuaranteeFee(conn, ts)
        '        'Case "RefundRecord"  '�����Ƿ����
        '        '    iCondition = New ImplEndReturn(conn, ts)
        '    Case Else  'һ������
        '        iCondition = New ImplCommon()

        'End Select

        'һ������
        iCondition = New ImplCommon

        Return iCondition.GetResult(workFlowID, projectID, taskID, expFlag, transCondition)

    End Function

    '���ù���
    Public Function VTask(ByVal workFlowID As String, ByVal ProjectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String)
        '1����������ȡ�����(ģ��ID����ĿID������ID)ƥ�������
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����  
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        Dim tmpTaskType As String = dsTempTask.Tables(0).Rows(0).Item("task_type")

        '2�������ǰ�����ṩ���̹��ߣ��������̹��ߣ�
        Dim tmpFlowTools As String
        Dim args As Object() = {conn, ts}
        If IsDBNull(dsTempTask.Tables(0).Rows(0).Item("flow_tool")) = False Then
            tmpFlowTools = dsTempTask.Tables(0).Rows(0).Item("flow_tool")
            tmpFlowTools = "BusinessRules." & tmpFlowTools

            '��̬�����ӿڶ���
            Dim t As System.Type = System.Type.GetType(tmpFlowTools)
            Dim iFlowTools As IFlowTools = Activator.CreateInstance(t, args)
            iFlowTools.UseFlowTools(workFlowID, ProjectID, taskID, finishedFlag, userID)

        End If


        '3����ȡת������
        Dim dsTempTaskTransfer As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '4�� ��ȡת������Ϊ���ת������
        Dim i, j, k As Integer
        Dim newRow As DataRow
        Dim nextTaskID, tmpTransCondition As String
        Dim dsConditionTrue As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")
        For i = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
            nextTaskID = dsTempTaskTransfer.Tables(0).Rows(i).Item("next_task")
            tmpTransCondition = Trim(dsTempTaskTransfer.Tables(0).Rows(i).Item("transfer_condition"))

            '�ж������Ƿ�Ϊ��
            If CompareExpression(workFlowID, ProjectID, taskID, ".T.", tmpTransCondition) Then

                '����ת������Ϊ���ת������
                newRow = dsConditionTrue.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = workFlowID
                    .Item("project_code") = ProjectID
                    .Item("task_id") = taskID
                    .Item("next_task") = nextTaskID
                    .Item("transfer_condition") = tmpTransCondition
                End With
                dsConditionTrue.Tables(0).Rows.Add(newRow)
            End If
        Next

        '5�������ǰ���������Ϊ�������������أ�
        If tmpTaskType = "END" Then
            Exit Function
        End If

        '6��[���������һ���к�̻������]��ÿ��ת������Ϊ���ת�����񣨶�����񣩣�
        Dim dsTempPreTask, dsTempPreTaskStatus As DataSet
        Dim mergeRelation, tmpPreTaskID, tmpNextTaskID, tmpTransStatus As String
        For i = 0 To dsConditionTrue.Tables(0).Rows.Count - 1
            tmpNextTaskID = dsConditionTrue.Tables(0).Rows(i).Item("next_task")
            strSql = "{project_code=" & "'" & ProjectID & "'" & " and task_id=" & "'" & tmpNextTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '�쳣����  
            If dsTempTask.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                Throw wfErr
            End If

            mergeRelation = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("merge_relation")), "", dsTempTask.Tables(0).Rows(0).Item("merge_relation"))

            tmpTransStatus = IIf(IsDBNull(dsConditionTrue.Tables(0).Rows(i).Item("project_status")), "", dsConditionTrue.Tables(0).Rows(i).Item("project_status"))

            If mergeRelation = "AND" Then
                '����㼯����Ϊ��AND������ȡת�������ǰ���״̬��

                strSql = "{project_code=" & "'" & ProjectID & "'" & " and next_task=" & "'" & tmpNextTaskID & "'" & "}"
                '��ȡת�������ǰ�����
                dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                '����ת�������ǰ�����
                For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1

                    tmpPreTaskID = dsTempTaskTransfer.Tables(0).Rows(j).Item("task_id")
                    strSql = "{project_code=" & "'" & ProjectID & "'" & " and task_id=" & "'" & tmpPreTaskID & "'" & "}"
                    '��ȡÿ��ǰ��������ɫ�������״̬��
                    dsTempPreTaskStatus = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    '����ǰ��������״̬���Ƿ�ȫΪ��F����
                    For k = 0 To dsTempPreTaskStatus.Tables(0).Rows.Count - 1
                        If dsTempPreTaskStatus.Tables(0).Rows(k).Item("task_status") <> "F" Then
                            Exit Function
                        End If
                    Next
                Next

                '����ת������
                StartupTask(workFlowID, ProjectID, tmpNextTaskID, "", "")

            Else

                '����ת������
                StartupTask(workFlowID, ProjectID, tmpNextTaskID, "", "")
            End If

        Next

    End Function

    '��ӹ�����־
    Public Function AddWorkLog(ByVal projectID As String, ByVal taskID As String, ByVal taskName As String, ByVal userID As String, ByVal taskStatus As String, ByVal startTime As DateTime, ByVal finishTime As DateTime, ByVal autoType As Integer, ByVal projectPhase As String, ByVal projectStatus As String, ByVal start_mode As String)
        Dim workLog As New WorkLog(conn, ts)
        Dim dsTempWorkLog As DataSet = workLog.GetWorkLogInfo("null")

        '��ȡ������Ľ�ɫID
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsAttendRole As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsAttendRole.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsAttendRole.Tables(0))
            Throw wfErr
        End If

        Dim roleID As String = Trim(dsAttendRole.Tables(0).Rows(0).Item("role_id"))

        Dim newRow As DataRow = dsTempWorkLog.Tables(0).NewRow
        With newRow
            .Item("project_code") = projectID
            .Item("task_id") = taskID
            .Item("task_name") = taskName
            .Item("role_id") = roleID
            .Item("attend_person") = userID
            .Item("task_status") = taskStatus
            .Item("start_time") = startTime
            .Item("finish_time") = finishTime
            .Item("auto") = autoType
            .Item("project_phase") = projectPhase
            .Item("project_status") = projectStatus
            .Item("start_mode") = start_mode
        End With
        dsTempWorkLog.Tables(0).Rows.Add(newRow)
        workLog.UpdateWorkLog(dsTempWorkLog)
    End Function

    'FiliterJeeDummyTask��ProjectID��ShiftTaskSet��
    '���˺�̻����������ȡҵ����
    Private Function FiliterJeeDummyTask(ByVal workflowID As String, ByVal projectID As String, ByVal ShiftTaskSet As DataSet, ByVal finishedFlag As String, ByVal userID As String) As DataSet

        '��	Vtask=True
        Dim vTask As Boolean = True

        Dim strSql As String
        Dim tmpTaskID, tmpTaskType, tmpApplyTool, tmpFlowTools, nextTaskID, tmpTransCondition, tmpTransStatus, tmpPhase, tmpStatus As String
        Dim dsTempNoVtaskSet, dsTempTask, dsTempTaskTransfer, dsConditionTrue, dsProject As DataSet
        Dim newRow As DataRow
        Dim args As Object() = {conn, ts}
        Dim i, j As Integer
        Dim t As System.Type
        Do While vTask = True

            vTask = False

            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsProject = project.GetProjectInfo(strSql)

            For i = 0 To ShiftTaskSet.Tables(0).Rows.Count - 1
                tmpStatus = Trim(IIf(IsDBNull(ShiftTaskSet.Tables(0).Rows(0).Item("project_status")), "", ShiftTaskSet.Tables(0).Rows(0).Item("project_status")))


                '��	����ת������Ϊ���ת�ƻ��ShiftTaskSet�е�ÿ������TaskID
                '    ��������ȡ��ǰ����ProjectID��TaskID����

                tmpTaskID = Trim(ShiftTaskSet.Tables(0).Rows(i).Item("next_task"))
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"


                '��ȡ����Ļ���ͣ�
                dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                '�쳣����  
                If dsTempTask.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                    Throw wfErr
                End If

                tmpTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))
                tmpPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))


                '�����������Ϊ��AUTO��
                If tmpTaskType = "AUTO" Then


                    If tmpPhase <> "" Then
                        dsProject.Tables(0).Rows(0).Item("phase") = tmpPhase
                    End If

                    If tmpStatus <> "" Then
                        dsProject.Tables(0).Rows(0).Item("status") = tmpStatus
                    End If

                    project.UpdateProject(dsProject)

                    'Vtask= True
                    vTask = True

                    '���Ӧ�ù��߷ǿգ�����Ӧ�ù��ߣ�
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                    dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                    '�쳣����  
                    If dsTempTask.Tables(0).Rows.Count = 0 Then
                        Dim wfErr As New WorkFlowErr
                        wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                        Throw wfErr
                    End If

                    'tmpApplyTool = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("apply_tool")), "", dsTempTask.Tables(0).Rows(0).Item("apply_tool")))

                    'If tmpApplyTool <> "" Then
                    '    '���ù���
                    '    tmpApplyTool = "BusinessRules." & tmpApplyTool
                    '    t = System.Type.GetType(tmpApplyTool)
                    '    Dim iApplyTools As IApplyTools = Activator.CreateInstance(t, args)
                    '    iApplyTools.UseApplyTools()

                    'End If

                    '�����ǰ�����ṩ���̹��ߣ��������̹��ߣ�
                    If Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("flow_tool")), "", dsTempTask.Tables(0).Rows(0).Item("flow_tool"))) <> "" Then

                        tmpFlowTools = dsTempTask.Tables(0).Rows(0).Item("flow_tool")
                        tmpFlowTools = "BusinessRules." & tmpFlowTools

                        '��̬�����ӿڶ���
                        t = System.Type.GetType(tmpFlowTools)
                        Dim iFlowTools As IFlowTools = Activator.CreateInstance(t, args)
                        iFlowTools.UseFlowTools(workflowID, projectID, tmpTaskID, finishedFlag, userID)

                    End If

                    '    ��ȡ��ǰ�����ת�������ת��������
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                    dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    '    ��ȡת������Ϊ���ת�����񼯣����ת������Ϊ�գ������棩��
                    dsConditionTrue = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")
                    For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
                        nextTaskID = dsTempTaskTransfer.Tables(0).Rows(j).Item("next_task")
                        tmpTransCondition = Trim(dsTempTaskTransfer.Tables(0).Rows(j).Item("transfer_condition"))
                        tmpTransStatus = Trim(IIf(IsDBNull(dsTempTaskTransfer.Tables(0).Rows(j).Item("project_status")), "", dsTempTaskTransfer.Tables(0).Rows(j).Item("project_status")))

                        '�ж������Ƿ�Ϊ��
                        If CompareExpression(workflowID, projectID, tmpTaskID, ".T.", tmpTransCondition) Then

                            '����ת������Ϊ���ת������
                            newRow = dsConditionTrue.Tables(0).NewRow()
                            With newRow
                                .Item("workflow_id") = ""
                                .Item("project_code") = projectID
                                .Item("task_id") = tmpTaskID
                                .Item("next_task") = nextTaskID
                                .Item("transfer_condition") = tmpTransCondition
                                .Item("project_status") = tmpTransStatus
                            End With
                            dsConditionTrue.Tables(0).Rows.Add(newRow)
                        End If
                    Next

                    '    ����ǰ������������ΪAUTO����ת�ƻ����ɾ����
                    ShiftTaskSet.Tables(0).Rows(i).Delete()
                    ShiftTaskSet.AcceptChanges()

                    '    ��ת������Ϊ���ת��������ӵ�ShiftTaskSet��
                    For j = 0 To dsConditionTrue.Tables(0).Rows.Count - 1
                        newRow = ShiftTaskSet.Tables(0).NewRow
                        With newRow
                            .Item("workflow_id") = ""
                            .Item("project_code") = projectID
                            .Item("task_id") = dsConditionTrue.Tables(0).Rows(j).Item("task_id")
                            .Item("next_task") = dsConditionTrue.Tables(0).Rows(j).Item("next_task")
                            .Item("transfer_condition") = dsConditionTrue.Tables(0).Rows(j).Item("transfer_condition")
                            .Item("project_status") = dsConditionTrue.Tables(0).Rows(j).Item("project_status")
                        End With
                        ShiftTaskSet.Tables(0).Rows.Add(newRow)
                    Next

                    '����ҵ�������,���ر��β���
                    Exit For

                End If


            Next
        Loop

        Return ShiftTaskSet

    End Function


    'FiliterRecedeDummyTask��ProjectID��ShiftTaskSet��
    '������Լ�����������ȡҵ��
    Private Function FiliterRecedeDummyTask(ByVal workflowID As String, ByVal projectID As String, ByVal ShiftTaskSet As DataSet, ByVal userID As String) As DataSet
        '��	Vtask=True
        Dim vTask As Boolean = True

        Dim strSql As String
        Dim tmpTaskID, tmpTaskType, tmpApplyTool, tmpFlowTools, nextTaskID, tmpTransCondition As String
        Dim dsTempNoVtaskSet, dsTempTask, dsTempTaskTransfer, dsConditionTrue As DataSet
        Dim newRow As DataRow
        Dim args As Object() = {conn, ts}
        Dim i, j As Integer
        Dim t As System.Type
        Do While vTask = True

            vTask = False

            For i = 0 To ShiftTaskSet.Tables(0).Rows.Count - 1
                '�� ����ת������Ϊ�����Լ���ShiftTaskSet�е�ÿ������TaskID
                '    ��������ȡ��ǰ����ProjectID��TaskID����
                tmpTaskID = ShiftTaskSet.Tables(0).Rows(i).Item("task_id")
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"

                '��ȡ����Ļ���ͣ�
                dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                '�쳣����  
                If dsTempTask.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                    Throw wfErr
                End If

                tmpTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))

                '�����������Ϊ��AUTO��
                If tmpTaskType = "AUTO" Then

                    'Vtask= True
                    vTask = True

                    '��ȡ��ǰ�������Լ�������Լ������
                    strSql = "{project_code=" & "'" & projectID & "'" & " and next_task=" & "'" & tmpTaskID & "'" & "}"
                    dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    '����ǰ������������ΪAUTO������Լ�����ɾ����
                    ShiftTaskSet.Tables(0).Rows(i).Delete()
                    ShiftTaskSet.AcceptChanges()

                    '����Լ������ӵ�ShiftTaskSet��
                    For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
                        newRow = ShiftTaskSet.Tables(0).NewRow
                        With newRow
                            .Item("workflow_id") = ""
                            .Item("project_code") = projectID
                            .Item("task_id") = dsTempTaskTransfer.Tables(0).Rows(j).Item("task_id")
                            .Item("next_task") = dsTempTaskTransfer.Tables(0).Rows(j).Item("next_task")
                            .Item("transfer_condition") = dsTempTaskTransfer.Tables(0).Rows(j).Item("transfer_condition")
                        End With
                        ShiftTaskSet.Tables(0).Rows.Add(newRow)
                    Next

                End If
            Next
        Loop

        Return ShiftTaskSet

    End Function


    'ˢ�»��飬�ѳ�����ǰϵͳʱ�仹û�а�����Ŀ�������ɾ��
    Public Function RefreshConference()
        ''��ȡ���г���ϵͳ��ǰʱ��Ļ���
        'Dim sysDate As String = FormatDateTime(Now, DateFormat.ShortDate)
        'Dim strSql As String
        'strSql = "{conference_date<" & "'" & sysDate & "'" & "}"
        'Dim Conference As New Conference(conn, ts)
        'Dim dsConference As DataSet = Conference.GetConferenceInfo(strSql, "null")

        ''ɾ��û�а�����Ŀ�������
        'Dim i As Integer
        'Dim tmpConferenceCode As String
        'Dim dsConfTrial As DataSet
        'Dim ConfTrial As New ConfTrial(conn, ts)
        'For i = 0 To dsConference.Tables(0).Rows.Count - 1
        '    tmpConferenceCode = dsConference.Tables(0).Rows(i).Item("conference_code")
        '    strSql = "{conference_code=" & "'" & tmpConferenceCode & "'" & "}"
        '    dsConfTrial = ConfTrial.GetConfTrialInfo(strSql, "null")

        '    '���û�а�����Ŀ
        '    If dsConfTrial.Tables(0).Rows.Count = 0 Then
        '        dsConference.Tables(0).Rows(i).Delete()
        '    End If

        'Next
        'Conference.UpdateConferenceCommitteeman(dsConference)

    End Function

    '�ύ���������������FinishedReviewConferencePlan��ConferenceCode��
    'Ϊ��֧����һ��������ϰ��Ŷ����Ŀ����������������ύ���������µ����� 
    Public Function FinishedReviewConferencePlan(ByVal ConferenceCode As String)
        '��	��Conference-Trail�л�ȡ����ConferenceCodeָ����������Ŀ���룻
        Dim strSql As String
        Dim i, j As Integer
        Dim newRow As DataRow
        Dim tmpProjectCode, tmpWorkflowID, tmpUserID, tmpManagerA, tmpManagerB, tmpTaskStatus As String
        Dim isHasA, isHasB As Boolean
        Dim dsTempProject, dsTempCommitteeman, dsTempAttend, dsTempMsg, dsConference As DataSet
        Dim ConfTrial As New ConfTrial(conn, ts)
        Dim Conference As New Conference(conn, ts)
        Dim CommonQuery As New CommonQuery(conn, ts)
        Dim record_person As String '�᳡�ļ�¼Ա

        Dim conferenceTime As DateTime = Now

        strSql = "{conference_code=" & "'" & ConferenceCode & "'" & "}"
        dsTempProject = ConfTrial.GetConfTrialInfo(strSql, "null")
        '��	��Conference-Committeeman���л�ȡ��������ί�����������˼���
        dsTempCommitteeman = Conference.GetConferenceInfo(strSql, strSql)

        If dsTempCommitteeman.Tables(0).Rows.Count > 0 Then
            Dim strTime As String
            strTime = dsTempCommitteeman.Tables(0).Rows(0).Item("conference_date") & " " & dsTempCommitteeman.Tables(0).Rows(0).Item("start_time") & ":00"
            conferenceTime = CDate(strTime)
        End If

        If dsTempCommitteeman.Tables("conference").Rows.Count > 0 Then
            Dim room As String = dsTempCommitteeman.Tables("conference").Rows(0)("place") & String.Empty
            Dim conferenceRoom As ConfernceRoom = New ConfernceRoom(conn, ts)
            Dim dsRoom As DataSet = conferenceRoom.FetchConfernceRoom(room)
            If dsRoom.Tables(0).Rows.Count > 0 Then
                record_person = dsRoom.Tables(0).Rows(0)("record_person") & String.Empty
            End If
            dsRoom.Dispose()
        End If

        Dim isExp As Boolean
        '��	��ÿ����Ŀ����
        For i = 0 To dsTempProject.Tables(0).Rows.Count - 1
            '��Project-Task-Attendee��ȡ����Ŀ����һ�£�ReviewMeetingPlan����״̬��P����Workflow-id�Ͳ����ˣ�
            tmpProjectCode = dsTempProject.Tables(0).Rows(i).Item("project_code")
            isExp = IIf(IsDBNull(dsTempProject.Tables(0).Rows(i).Item("is_exp")), False, dsTempProject.Tables(0).Rows(i).Item("is_exp"))

            '�����չ����Ŀ���ύչ�ڵð��Ż�������
            If isExp Then
                strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlanExp'}"
            Else
                strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan'}"
            End If

            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpTaskStatus = dsTempAttend.Tables(0).Rows(0).Item("task_status")

            '�ж���Ŀ�Ƿ����ύ��,ֻ��δ�ύ������Ŀ�ſɰ���
            If tmpTaskStatus = "P" Then


                tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")
                tmpUserID = Trim(dsTempAttend.Tables(0).Rows(0).Item("attend_person"))

                '��ȡ��Workflow-id����Ŀ����ƥ�����Ŀ����A�Ǻ�B�ǣ�
                '�����Ŀ����A�ǲ��ڻ�������˼��У�����Ŀ����A�Ǽ����������˼���
                strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
                dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

                '�쳣����  
                If dsTempAttend.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                    Throw wfErr
                End If

                tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))

                '�����Ŀ����B�ǲ��ڻ�������˼��У�����Ŀ����B�Ǽ����������˼���
                tmpManagerB = Trim(dsTempAttend.Tables(0).Rows(0).Item("25"))

                For j = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
                    If tmpManagerA = dsTempCommitteeman.Tables(1).Rows(j).Item("committeeman") Then
                        isHasA = True
                        Exit For
                    End If
                Next

                For j = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
                    If tmpManagerB = dsTempCommitteeman.Tables(1).Rows(j).Item("committeeman") Then
                        isHasB = True
                        Exit For
                    End If
                Next

                If isHasA = False Then
                    newRow = dsTempCommitteeman.Tables(1).NewRow
                    newRow.Item("conference_code") = ConferenceCode
                    newRow.Item("committeeman") = tmpManagerA
                    dsTempCommitteeman.Tables(1).Rows.Add(newRow)
                End If

                If isHasB = False Then
                    newRow = dsTempCommitteeman.Tables(1).NewRow
                    newRow.Item("conference_code") = ConferenceCode
                    newRow.Item("committeeman") = tmpManagerB
                    dsTempCommitteeman.Tables(1).Rows.Add(newRow)
                End If

                '���Ĵ��� ��¼�������� �������ԱΪ �᳡�ļ�¼��Ա  2005-6-30 LQF add
                If isExp Then
                    strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusionExp'}"
                Else
                    strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'}"
                End If

                dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                If dsTempAttend.Tables(0).Rows.Count <> 0 Then
                    dsTempAttend.Tables(0).Rows(0)("attend_person") = record_person
                    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)
                End If

                '�ж���Ŀ�Ƿ����ύ

                '�����չ����Ŀ���ύչ�ڵð��Ż�������
                If isExp Then
                    strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlanExp' and task_status='P'}"
                Else
                    strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan' and task_status='P'}"
                End If

                dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                If dsTempAttend.Tables(0).Rows.Count <> 0 Then
                    If isExp Then
                        '����FinishedTask��Workflow-id����Ŀ���롢ReviewMeetingPlan���� ���������ˣ���
                        finishedTask(tmpWorkflowID, tmpProjectCode, "ReviewMeetingPlanExp", "", tmpUserID)
                        setReviewConclusionCueTime(tmpProjectCode, "ReviewMeetingPlanExp", conferenceTime)
                    Else
                        '����FinishedTask��Workflow-id����Ŀ���롢ReviewMeetingPlan���� ���������ˣ���
                        finishedTask(tmpWorkflowID, tmpProjectCode, "ReviewMeetingPlan", "", tmpUserID)
                        setReviewConclusionCueTime(tmpProjectCode, "ReviewMeetingPlan", conferenceTime)
                    End If


                End If
            End If

        Next

        ''��	���ڻ�������˼������г�Ա
        ''����Ϣ����ӡ���鿴�������̡���Ϣ��
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")


        ''��ȡ����������
        'strSql = "{conference_code=" & "'" & ConferenceCode & "'" & "}"
        'dsConference = Conference.GetConferenceInfo(strSql, "null")

        ''�쳣����  
        'If dsConference.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsConference.Tables(0))
        '    Throw wfErr
        'End If

        'Dim tmpConfTime As String = CStr(dsConference.Tables(0).Rows(0).Item("conference_date")) & " " & dsConference.Tables(0).Rows(0).Item("start_time")

        'For i = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = "��鿴" & tmpConfTime & "������������"
        '        .Item("accepter") = dsTempCommitteeman.Tables(1).Rows(i).Item("committeeman")
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)

    End Function

    '����������鰲������CancelReviewConferencePlan��ConferenceCode��
    Public Function CancelReviewConferencePlan(ByVal ConferenceCode As String)
        '��	��Conference-Trail�л�ȡ����ConferenceCodeָ����������Ŀ���룻
        Dim strSql As String
        Dim i, j As Integer
        Dim newRow As DataRow
        Dim tmpProjectCode, tmpWorkflowID, tmpTaskID, tmpUserID, tmpManagerA, tmpManagerB As String
        Dim tmpConferenceDate As DateTime
        Dim isHasA, isHasB As Boolean
        Dim dsTempProject, dsTempConference, dsTempCommitteeman, dsTempAttend, dsTempMsg As DataSet
        Dim ConfTrial As New ConfTrial(conn, ts)
        Dim Conference As New Conference(conn, ts)
        Dim CommonQuery As New CommonQuery(conn, ts)

        strSql = "{conference_code=" & "'" & ConferenceCode & "'" & "}"
        dsTempProject = ConfTrial.GetConfTrialInfo(strSql, "null")

        '��	��Conference-Committeeman���ȡ��������ί�����������˼���
        dsTempCommitteeman = Conference.GetConferenceInfo("null", strSql)

        '��	��Conference���л�ȡ��ConferenceCodeƥ���Conference-date��
        dsTempConference = Conference.GetConferenceInfo(strSql, "null")

        '�쳣����  
        If dsTempConference.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempConference.Tables(0))
            Throw wfErr
        End If

        tmpConferenceDate = dsTempConference.Tables(0).Rows(0).Item("conference_date")
        '��	��ÿ����Ŀ����
        For i = 0 To dsTempProject.Tables(0).Rows.Count - 1
            tmpProjectCode = dsTempProject.Tables(0).Rows(i).Item("project_code")

            '��ȡ����Ŀ��workflow_id
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")

            '�ж�������Ƿ��ѿ���
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion' and task_status='F'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            If dsTempAttend.Tables(0).Rows.Count > 0 Then
                '�׳�"���ܳ����ѿ����������"
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowRecordReviewConclusionErr()
                Throw wfErr
                Exit Function
            End If

            '��������Ŀ�ļ�¼�������������ÿ�
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For j = 0 To dsTempAttend.Tables(0).Rows.Count - 1
                dsTempAttend.Tables(0).Rows(j).Item("task_status") = ""
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            ''��Project-Task-Attendee��ȡ����Ŀ����һ�£�ReviewMeetingPlan����״̬Ϊ��F��������
            'strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan' and task_status='F'" & "}"
            'dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            ''�쳣����  
            'If dsTempAttend.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
            '    Throw wfErr
            'End If

            'tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")
            'tmpTaskID = Trim(dsTempAttend.Tables(0).Rows(0).Item("task_id"))

            '��ȡ��Workflow-id����Ŀ����ƥ�����Ŀ����A�Ǻ�B�ǣ�
            '�����Ŀ����A�ǲ��ڻ�������˼��У�����Ŀ����A�Ǽ����������˼���
            strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
            dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))

            '�����Ŀ����B�ǲ��ڻ�������˼��У�����Ŀ����B�Ǽ����������˼���
            tmpManagerB = Trim(dsTempAttend.Tables(0).Rows(0).Item("25"))

            For j = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
                If tmpManagerA = dsTempCommitteeman.Tables(1).Rows(j).Item("committeeman") Then
                    isHasA = True
                    Exit For
                End If
            Next

            For j = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
                If tmpManagerB = dsTempCommitteeman.Tables(1).Rows(j).Item("committeeman") Then
                    isHasB = True
                    Exit For
                End If
            Next

            If isHasA Then
                newRow = dsTempCommitteeman.Tables(1).NewRow
                newRow.Item("conference_code") = ConferenceCode
                newRow.Item("committeeman") = tmpManagerA
                dsTempCommitteeman.Tables(1).Rows.Add(newRow)
            End If

            If isHasB Then
                newRow = dsTempCommitteeman.Tables(1).NewRow
                newRow.Item("conference_code") = ConferenceCode
                newRow.Item("committeeman") = tmpManagerB
                dsTempCommitteeman.Tables(1).Rows.Add(newRow)
            End If

            '������״̬��Ϊ��P����
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For j = 0 To dsTempAttend.Tables(0).Rows.Count - 1
                dsTempAttend.Tables(0).Rows(j).Item("task_status") = "P"
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            ''�����¼��������״̬Ϊ'P' ,������״̬��Ϊ������
            'strSql = "{workflow_id=" & "'" & tmpWorkflowID & "'" & " and project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'" & "}"
            'dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            'If IIf(IsDBNull(dsTempAttend.Tables(0).Rows(0).Item("task_status")), "", dsTempAttend.Tables(0).Rows(0).Item("task_status")) = "P" Then
            '    dsTempAttend.Tables(0).Rows(0).Item("task_status") = ""
            'End If
            'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        Next

        ''��	���ڻ�������˼������г�Ա
        ''����Ϣ����ӡ�Conference-date����᳷������Ϣ��
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")
        'For i = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = CStr(tmpConferenceDate) & "����᳷��"
        '        .Item("accepter") = dsTempCommitteeman.Tables(1).Rows(i).Item("committeeman")
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)

        'ɾ���û����������Ŀ
        For i = 0 To dsTempProject.Tables(0).Rows.Count - 1
            dsTempProject.Tables(0).Rows(i).Item("conference_code") = DBNull.Value
        Next
        ConfTrial.UpdateConfTrial(dsTempProject)

        '����ɾ��
        dsTempConference.Tables(0).Rows(0).Delete()
        Conference.UpdateConferenceCommitteeman(dsTempConference)
    End Function

    '�����������Ŀ
    Public Function CancelReviewConferencePlanProject(ByVal projectID As String)
        Dim strSql As String
        Dim dsTempAttend As DataSet
        Dim i As Integer

        '�ڲ����˱��н�����Ŀ�ļ�¼��������������Ϊ""
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordReviewConclusion'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '�ڲ����˱��н�����Ŀ�İ��������������Ϊ"P"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ReviewMeetingPlan'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function

    '�Զ�ȷ����ϢAACKMassage��ProjectID��TaskID��EmployeeID��
    Public Function AACKMassage(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String, ByVal employeeID As String)

        '��	��ȡProjectID����ҵ���ƣ�
        '��ȡ��Ŀ����ҵ����
        Dim strSql As String
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = project.GetProjectInfo(strSql)

        '�쳣����  
        If dsProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProject.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")

        '�쳣����  
        If dsCorporation.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsCorporation.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '��	��ȡProjectID��TaskID���������ƣ�
        '��	��ȡ����ID���������ƣ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        'δ�ҵ������򷵻�  
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Exit Function
        Else

            Dim tmpTaskName As String = dsTempTask.Tables(0).Rows(0).Item("task_name")

            '��	����Ϣ������Ϣ���ݰ�����ҵ���ƺ��������Ƶ���Ϣȷ�ϱ�־��Ϊ��Y����
            strSql = "{is_affirmed<>'Y'and project_code=" & "'" & projectID & "'" & "}"
            Dim dsTempMsg As DataSet = WfProjectMessages.GetWfProjectMessagesInfo(strSql)
            Dim tmpMsgContent, tmpAccepter As String
            Dim i As Integer
            Dim iCorp, iTask As Integer
            For i = 0 To dsTempMsg.Tables(0).Rows.Count - 1
                tmpMsgContent = IIf(IsDBNull(dsTempMsg.Tables(0).Rows(i).Item("message_content")), "", dsTempMsg.Tables(0).Rows(i).Item("message_content"))
                tmpAccepter = IIf(IsDBNull(dsTempMsg.Tables(0).Rows(i).Item("accepter")), "", dsTempMsg.Tables(0).Rows(i).Item("accepter"))
                iCorp = InStr(tmpMsgContent, tmpCorporationName)
                iTask = InStr(tmpMsgContent, tmpTaskName)
                If iCorp <> 0 And iTask <> 0 Then
                    dsTempMsg.Tables(0).Rows(i).Item("is_affirmed") = "Y"
                End If
            Next
            WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)
        End If

    End Function

    '��ί�У���ȡ���������getTaskActor��RoleID��
    Public Function getTaskActor(ByVal roleID As String) As String
        '��	��STAFF-ROLE���ȡ��ɫID=RoleID��staff_name��consigner;
        Dim strSql As String
        Dim i As Integer
        Dim role As New Role(conn, ts)
        strSql = "{role_id=" & "'" & roleID & "'" & "}"
        Dim dsTempStaff As DataSet = role.GetStaffRole(strSql)

        '�쳣����  
        If dsTempStaff.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
            Throw wfErr
        End If

        '��	Staff_name=staff_name��
        Dim tmpStaffName As String = Trim(dsTempStaff.Tables(0).Rows(0).Item("staff_name"))
        '��	Consigner=consigner��
        Dim tmpConsigner As String = Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(0).Item("consigner")), "", dsTempStaff.Tables(0).Rows(0).Item("consigner")))
        '��	WHILE ��Consigner�ǿգ�
        While tmpConsigner <> ""
            '��STAFF-ROLE���ȡԱ������ΪConsigner��staff_name��ί���˼�consignerSet��
            strSql = "{staff_name=" & "'" & tmpConsigner & "'" & "}"
            dsTempStaff = role.GetStaffRole(strSql)

            '�쳣����  
            If dsTempStaff.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
                Throw wfErr
            End If

            'Staff_name=Consigner��
            tmpStaffName = tmpConsigner
            'Consigner=""
            tmpConsigner = ""
            '   For I = 0 To Number(consignerSet) - 1
            For i = 0 To dsTempStaff.Tables(0).Rows.Count - 1
                'IF consignerSet��I��IS NOT Null THEN Consigner=consigner��
                If Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(0).Item("consigner")), "", dsTempStaff.Tables(0).Rows(0).Item("consigner"))) <> "" Then
                    tmpConsigner = dsTempStaff.Tables(0).Rows(0).Item("consigner")
                End If
            Next
        End While

        '��	����Staff_name��
        Return tmpStaffName
    End Function

    Public Function getTaskActor(ByVal projectID As String, ByVal taskID As String, ByVal roleID As String, ByVal branch As String) As String
        '��	��STAFF-ROLE���ȡ��ɫID=RoleID��staff_name��consigner;
        Dim strSql As String
        Dim i As Integer
        Dim role As New Role(conn, ts)
        strSql = "{role_id=" & "'" & roleID & "'" & "}"
        Dim dsTempStaff As DataSet = role.GetStaffRole(strSql)

        '�쳣����  
        If dsTempStaff.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
            Throw wfErr
        End If

        Dim tmpStaffName As String
        Dim dsTemp As DataSet
        Dim staff As New Staff(conn, ts)
        Dim isFound As Boolean
        Dim iStaff As Integer
        For i = 0 To dsTempStaff.Tables(0).Rows.Count - 1
            '��	Staff_name=staff_name��
            tmpStaffName = Trim(dsTempStaff.Tables(0).Rows(i).Item("staff_name"))
            strSql = "{staff_name=" & "'" & tmpStaffName & "'" & "}"
            dsTemp = staff.FetchStaff(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            If IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("branch_name")), "", dsTemp.Tables(0).Rows(0).Item("branch_name")) = branch Then
                isFound = True
                iStaff = i
                Exit For
            End If
        Next

        '����ڷ�֧�����ҵ�������
        If isFound Then

            '��	Consigner=consigner��
            Dim tmpConsigner As String = Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(iStaff).Item("consigner")), "", dsTempStaff.Tables(0).Rows(iStaff).Item("consigner")))
            '��	WHILE ��Consigner�ǿգ�
            'While tmpConsigner <> ""
            '    '��STAFF-ROLE���ȡԱ������ΪConsigner��staff_name��ί���˼�consignerSet��
            '    strSql = "{staff_name=" & "'" & tmpConsigner & "'" & "}"
            '    dsTempStaff = role.GetStaffRole(strSql)

            '    '�쳣����  
            '    If dsTempStaff.Tables(0).Rows.Count = 0 Then
            '        Dim wfErr As New WorkFlowErr()
            '        wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
            '        Throw wfErr
            '    End If

            '    'Staff_name=Consigner��
            '    tmpStaffName = tmpConsigner
            '    'Consigner=""
            '    tmpConsigner = ""
            '    '   For I = 0 To Number(consignerSet) - 1
            '    For i = 0 To dsTempStaff.Tables(0).Rows.Count - 1
            '        'IF consignerSet��I��IS NOT Null THEN Consigner=consigner��
            '        If Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(i).Item("consigner")), "", dsTempStaff.Tables(0).Rows(i).Item("consigner"))) <> "" Then
            '            tmpConsigner = dsTempStaff.Tables(0).Rows(0).Item("consigner")
            '        End If
            '    Next
            'End While

            If tmpConsigner <> "" Then
                tmpStaffName = tmpConsigner

                '��ԭί��������consinger��
                Dim tmpSrcPerson As String = Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(iStaff).Item("staff_name")), "", dsTempStaff.Tables(0).Rows(iStaff).Item("staff_name")))
                Dim dsConsinger As DataSet
                strSql = "{project_code='" & projectID & "' and task_id='" & taskID & "'}"
                dsConsinger = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                    dsConsinger.Tables(0).Rows(i).Item("consigner") = tmpSrcPerson
                Next

                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsConsinger)

            End If


            '��	����Staff_name��
            Return tmpStaffName
        Else

            Return ""

        End If

    End Function


    'consignTask��Ա������ɫ��ί���ˣ����ɿͻ���ʹ�á�
    Public Function consignTask(ByVal staffID As String, ByVal roleID As String, ByVal consigner As String, ByVal isCurrent As Boolean)
        '��	��STAFF-ROLE����Ҳ���ָ����Ա������ɫ��
        Dim strSql As String
        strSql = "{staff_name=" & "'" & staffID & "'" & " and role_id=" & "'" & roleID & "'" & "}"
        Dim role As New Role(conn, ts)
        Dim dsTempRole As DataSet = role.GetStaffRole(strSql)


        '�쳣����  
        If dsTempRole.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
            Throw wfErr
        End If

        ''�����Ա����ί����������ʾ������ί�д���
        'Dim tmpConsigner As String = IIf(IsDBNull(dsTempRole.Tables(0).Rows(0).Item("consigner")), "", dsTempRole.Tables(0).Rows(0).Item("consigner"))

        'If tmpConsigner <> "" Then
        '    '��ʾ������ί�����������
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowIsConsign()
        '    Throw wfErr
        '    Exit Function
        'End If

        '��	���ָ����Ա����ɫ����
        If dsTempRole.Tables(0).Rows.Count <> 0 Then
            '��ROLE���ȡԱ��ί�н�ɫ��ί�б�־��
            strSql = "{role_id=" & "'" & roleID & "'" & "}"
            Dim dsTempRoleConsign As DataSet = role.FetchRole(strSql)

            '�쳣����  
            If dsTempRoleConsign.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempRoleConsign.Tables(0))
                Throw wfErr
            End If


            '���Isconsign=1 
            If IIf(IsDBNull(dsTempRoleConsign.Tables(0).Rows(0).Item("isConsign")), False, dsTempRoleConsign.Tables(0).Rows(0).Item("isConsign")) = True Then

                ' ��STAFF-ROLE���е�Ա��ί����consigner��Ϊί���ˣ�
                dsTempRole.Tables(0).Rows(0).Item("consigner") = consigner
                role.UpdateStaffRole(dsTempRole)

                '���Ҫί�е�ǰ����,����consignCurrentTask
                'If isCurrent Then
                consignCurrentTask(staffID, roleID, consigner)
                'End If

                '����
            Else

                '��ʾ��û��ί��Ȩ�ޣ���
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoConsignRight()
                Throw wfErr

            End If
        End If
    End Function


    'CancelconsignTask��Ա������ɫ�����ɿͻ���ʹ�á�
    Public Function CancelconsignTask(ByVal srcPerson As String, ByVal staffID As String, ByVal roleID As String, ByVal isCurrent As Boolean)
        '��	��STAFF-ROLE����������Ա������ɫƥ��ļ�¼��
        Dim strSql As String
        strSql = "{staff_name=" & "'" & srcPerson & "'" & " and role_id=" & "'" & roleID & "'" & "}"
        Dim role As New Role(conn, ts)
        Dim dsTempRole As DataSet = role.GetStaffRole(strSql)

        '�쳣����  
        If dsTempRole.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
            Throw wfErr
        End If


        '��	������Ҽ�¼���ڣ���consigner��ΪNULL��
        If dsTempRole.Tables(0).Rows.Count <> 0 Then
            If dsTempRole.Tables(0).Rows(0).Item("consigner") Is DBNull.Value Then
                '��ʾ����������!��
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoConsigner()
                Throw wfErr
            Else
                dsTempRole.Tables(0).Rows(0).Item("consigner") = DBNull.Value

                '���Ҫ������ǰ�����ί��,����CancelconsignCurrentTask
                'If isCurrent Then
                CancelconsignCurrentTask(staffID, roleID, srcPerson)
                ' End If

            End If
        End If

        role.UpdateStaffRole(dsTempRole)

    End Function

    'ί�е�ǰ����
    Private Function consignCurrentTask(ByVal staffID As String, ByVal roleID As String, ByVal consigner As String)
        ''�������˱�������ΪP��ROLEID=ί�н�ɫ����������˸�Ϊ���ղ�����
        'Dim strSql As String
        'Dim i As Integer
        'strSql = "{role_id='" & roleID & "' and attend_person='" & staffID & "' and task_status='P'}"
        'Dim dsTempAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        ''�����ǰ����Ϊ�ʲ�����CapitialEvaluated���轫�������������ʦ��Ϊ������
        'strSql = "{project_code IN " & _
        '        "(SELECT DISTINCT project_code FROM project_task_attendee " & _
        '         "WHERE role_id='" & roleID & "' AND attend_person='" & staffID & _
        '         "' AND task_status='P' AND task_id='CapitialEvaluated') " & _
        '         "AND evaluate_person='" & staffID & "'}"
        'Dim dsOppositeGuarantee As DataSet
        'Dim ObjOppositeGuarantee As New Guaranty(conn, ts)
        'dsOppositeGuarantee = ObjOppositeGuarantee.GetGuarantyInfo(strSql, "null")
        'For i = 0 To dsOppositeGuarantee.Tables(0).Rows.Count - 1
        '    dsOppositeGuarantee.Tables(0).Rows(i).Item("evaluate_person") = consigner
        'Next
        'ObjOppositeGuarantee.UpdateGuaranty(dsOppositeGuarantee)

        ''��¼ԭί����
        '' ��ȡԭ������
        '' ��¼ԭ������
        'Dim dsSrcPerson As String
        'For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
        '    dsSrcPerson = dsTempAttend.Tables(0).Rows(i).Item("attend_person")
        '    dsTempAttend.Tables(0).Rows(i).Item("consigner") = dsSrcPerson
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        ''�������˸�Ϊ������
        ''Dim tmpLastAttend As String = getTaskActor(roleID)
        'For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
        '    dsTempAttend.Tables(0).Rows(i).Item("attend_person") = consigner
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '�������˱�������ΪP��ROLEID=ί�н�ɫ����������˸�Ϊ���ղ�����
        Dim strSql As String
        Dim i As Integer
        strSql = "{role_id='" & roleID & "' and attend_person='" & staffID & "'}"
        Dim dsTempAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        '�����ǰ����Ϊ�ʲ�����CapitialEvaluated���轫�������������ʦ��Ϊ������
        strSql = "{project_code IN " & _
                "(SELECT DISTINCT project_code FROM project_task_attendee " & _
                 "WHERE role_id='" & roleID & "' AND attend_person='" & staffID & _
                 "' AND task_status='P' AND task_id='CapitialEvaluated') " & _
                 "AND evaluate_person='" & staffID & "'}"
        Dim dsOppositeGuarantee As DataSet
        Dim ObjOppositeGuarantee As New Guaranty(conn, ts)
        dsOppositeGuarantee = ObjOppositeGuarantee.GetGuarantyInfo(strSql, "null")
        For i = 0 To dsOppositeGuarantee.Tables(0).Rows.Count - 1
            dsOppositeGuarantee.Tables(0).Rows(i).Item("evaluate_person") = consigner
        Next
        ObjOppositeGuarantee.UpdateGuaranty(dsOppositeGuarantee)

        '��¼ԭί����
        ' ��ȡԭ������
        ' ��¼ԭ������
        '�������˸�Ϊ������
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("attend_person") = consigner
            dsTempAttend.Tables(0).Rows(i).Item("consigner") = staffID
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function

    '������ǰ����ί��
    Private Function CancelconsignCurrentTask(ByVal staffID As String, ByVal roleID As String, ByVal srcPerson As String)
        ''�������˱�������ΪP��ROLEID=ί�н�ɫ����������˸�Ϊԭ��ɫԱ��
        'Dim strSql As String
        'Dim i As Integer
        'Dim dsTempAttend As DataSet

        ''�����ǰ����Ϊ�ʲ�����CapitialEvaluated���轫�������������ʦ��Ϊԭί����
        'strSql = "{project_code IN " & _
        '     "(SELECT DISTINCT project_code FROM project_task_attendee " & _
        '      "WHERE role_id='" & roleID & "' AND attend_person='" & staffID & _
        '      "' AND task_status='P' AND task_id='CapitialEvaluated') " & _
        '      "AND evaluate_person='" & staffID & "'}"
        'Dim dsOppositeGuarantee As DataSet
        'Dim ObjOppositeGuarantee As New Guaranty(conn, ts)
        'dsOppositeGuarantee = ObjOppositeGuarantee.GetGuarantyInfo(strSql, "null")
        'For i = 0 To dsOppositeGuarantee.Tables(0).Rows.Count - 1
        '    dsOppositeGuarantee.Tables(0).Rows(i).Item("evaluate_person") = srcPerson
        'Next
        'ObjOppositeGuarantee.UpdateGuaranty(dsOppositeGuarantee)


        'strSql = "{role_id='" & roleID & "' and attend_person='" & staffID & "' and task_status='P'}"
        'dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        'Dim dsSrcPerson As String
        ''�������˸�Ϊԭί����
        'For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1

        '    '��ȡ�������ԭί����
        '    dsSrcPerson = IIf(IsDBNull(dsTempAttend.Tables(0).Rows(i).Item("consigner")), "", dsTempAttend.Tables(0).Rows(i).Item("consigner"))

        '    '���ԭί���˷ǿ�
        '    If dsSrcPerson <> "" Then
        '        dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsSrcPerson
        '        dsTempAttend.Tables(0).Rows(i).Item("consigner") = ""
        '    End If
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '�������˱�������ΪP��ROLEID=ί�н�ɫ����������˸�Ϊԭ��ɫԱ��
        Dim strSql As String
        Dim i As Integer
        Dim dsTempAttend As DataSet

        '�����ǰ����Ϊ�ʲ�����CapitialEvaluated���轫�������������ʦ��Ϊԭί����
        strSql = "{project_code IN " & _
             "(SELECT DISTINCT project_code FROM project_task_attendee " & _
              "WHERE role_id='" & roleID & "' AND attend_person='" & staffID & _
              "' AND task_status='P' AND task_id='CapitialEvaluated') " & _
              "AND evaluate_person='" & staffID & "'}"
        Dim dsOppositeGuarantee As DataSet
        Dim ObjOppositeGuarantee As New Guaranty(conn, ts)
        dsOppositeGuarantee = ObjOppositeGuarantee.GetGuarantyInfo(strSql, "null")
        For i = 0 To dsOppositeGuarantee.Tables(0).Rows.Count - 1
            dsOppositeGuarantee.Tables(0).Rows(i).Item("evaluate_person") = srcPerson
        Next
        ObjOppositeGuarantee.UpdateGuaranty(dsOppositeGuarantee)


        strSql = "{role_id='" & roleID & "' and attend_person='" & staffID & "'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        Dim dsSrcPerson As String
        '�������˸�Ϊԭί����
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1

            '��ȡ�������ԭί����
            dsSrcPerson = IIf(IsDBNull(dsTempAttend.Tables(0).Rows(i).Item("consigner")), "", dsTempAttend.Tables(0).Rows(i).Item("consigner"))

            '���ԭί���˷ǿ�
            If dsSrcPerson <> "" Then
                dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsSrcPerson
                dsTempAttend.Tables(0).Rows(i).Item("consigner") = ""
            End If
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function


    '��ʽ��AddTaskTrackRecord��ProjectID,Workflow_id,TaskID,StartupTask��
    Public Function AddTaskTrackRecord(ByVal workflowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal StartupTaskID As String)
        Dim strSql As String
        Dim i As Integer

        '��	����Project_Track����;
        Dim dsProjectTrack As DataSet
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo("null")

        '��	������ֵ�ֱ������������Project_Code��Workflow_id��FinishedTask��StartupTask����,Status������Ϊ��P������
        Dim newRow As DataRow = dsProjectTrack.Tables(0).NewRow
        With newRow
            .Item("project_code") = projectID
            .Item("workflow_id") = workflowID
            .Item("FinishedTask") = finishedTaskID
            .Item("StartupTask") = StartupTaskID
            .Item("Status") = "P"
        End With
        dsProjectTrack.Tables(0).Rows.Add(newRow)
        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

        '��	����������Project_Code= ProjectID ��StartupTask= TaskID��Status=��P����Project_Track�����Status������Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & finishedTaskID & "'" & " and isnull(status,'')='P'}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)
        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1
            dsProjectTrack.Tables(0).Rows(i).Item("Status") = "F"
        Next
        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

    End Function


    '�ύǩԼ�ƻ�
    Public Function FinishedSignaturePlan(ByVal SignaturePlanCode As Integer)
        Dim i, j As Integer
        Dim strSql As String
        Dim ProjectSignature As New ProjectSignature(conn, ts)
        Dim SignaturePlan As New SignaturePlan(conn, ts)
        Dim dsProjectSignature, dsTempAttend, dsTempMsg, dsSignaturePlan, dsTimingTask As DataSet
        Dim tmpProjectCode, tmpWorkflowID, tmpUserID, tmpTaskStatus, tmpManagerA, tmpManagerB, tmpMinister, tmpManagerLaw, tmpDirector As String
        Dim CommonQuery As New CommonQuery(conn, ts)

        '��ȡ��ǩԼ�ƻ���������Ŀ
        strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        dsProjectSignature = ProjectSignature.GetProjectSignatureInfo(strSql)

        strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        dsSignaturePlan = SignaturePlan.GetSignaturePlanInfo(strSql)

        '�쳣����  
        If dsSignaturePlan.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsSignaturePlan.Tables(0))
            Throw wfErr
        End If

        Dim tmpSignaturePlanDate As DateTime = dsSignaturePlan.Tables(0).Rows(0).Item("signature_plan_date")


        For i = 0 To dsProjectSignature.Tables(0).Rows.Count - 1
            tmpProjectCode = dsProjectSignature.Tables(0).Rows(i).Item("project_code")

            '��Project-Task-Attendee��ȡ����Ŀ����һ�£�PlanSignature����״̬Ϊ��P����Workflow-id�Ͳ����ˣ�
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='PlanSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpTaskStatus = dsTempAttend.Tables(0).Rows(0).Item("task_status")

            '�ж���Ŀ�Ƿ����ύ��,ֻ��δ�ύ������Ŀ�ſɰ���
            If tmpTaskStatus = "P" Then


                tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")
                tmpUserID = Trim(dsTempAttend.Tables(0).Rows(0).Item("attend_person"))

                '����FinishedTask��Workflow-id����Ŀ���롢PlanSignature���� ���������ˣ���
                finishedTask(tmpWorkflowID, tmpProjectCode, "PlanSignature", "", tmpUserID)

                '���Ǽ�ǩԼ��ʱ�����״̬��Ϊ"P",��ʼʱ����ΪǩԼ�ƻ���ʱ��
                strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'}"
                dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                For j = 0 To dsTimingTask.Tables(0).Rows.Count - 1
                    dsTimingTask.Tables(0).Rows(j).Item("status") = "P"
                    dsTimingTask.Tables(0).Rows(j).Item("start_time") = tmpSignaturePlanDate
                Next
                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

            End If
        Next

        '��ȡ��Workflow-id����Ŀ����ƥ�����Ŀ����A�Ǻ�B��,�������Σ����ղ��������ﾭ��
        '�����Ŀ����A�ǲ��ڻ�������˼��У�����Ŀ����A�Ǽ����������˼���
        strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
        dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

        '�쳣����  
        If dsTempAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
            Throw wfErr
        End If

        Dim ArrAttend As New ArrayList
        tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))
        ArrAttend.Add(tmpManagerA)

        '�����Ŀ����B�ǲ��ڻ�������˼��У�����Ŀ����B�Ǽ����������˼���
        tmpManagerB = Trim(dsTempAttend.Tables(0).Rows(0).Item("25"))
        ArrAttend.Add(tmpManagerB)
        tmpMinister = Trim(dsTempAttend.Tables(0).Rows(0).Item("31"))
        ArrAttend.Add(tmpMinister)
        tmpManagerLaw = Trim(dsTempAttend.Tables(0).Rows(0).Item("33"))

        '���ղ����ͷ��ﾭ��Ϊͬһ��
        If tmpManagerLaw <> tmpMinister Then
            ArrAttend.Add(tmpManagerLaw)
        End If

        tmpDirector = getTaskActor("01")
        ArrAttend.Add(tmpDirector)

        ''��	����ǩԼ�����˼������г�Ա
        ''����Ϣ����ӡ�Signature-plan-date����ϯǩԼ��Ϣ��
        'Dim newRow As DataRow
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")
        'For i = 0 To ArrAttend.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = CStr(tmpSignaturePlanDate) & "���ϯǩԼ"
        '        .Item("accepter") = ArrAttend(i)
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)


    End Function

    '����ǩԼ
    Public Function CancelSignaturePlan(ByVal SignaturePlanCode As Integer)
        Dim i, j As Integer
        Dim strSql As String
        Dim ProjectSignature As New ProjectSignature(conn, ts)
        Dim SignaturePlan As New SignaturePlan(conn, ts)
        Dim dsProjectSignature, dsTempAttend, dsTempMsg, dsSignaturePlan, dsTimingTask As DataSet
        Dim tmpProjectCode, tmpWorkflowID, tmpUserID, tmpTaskID, tmpManagerA, tmpManagerB, tmpMinister, tmpManagerLaw, tmpDirector As String
        Dim CommonQuery As New CommonQuery(conn, ts)

        '��ȡ��ǩԼ�ƻ���������Ŀ
        strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        dsProjectSignature = ProjectSignature.GetProjectSignatureInfo(strSql)

        '��	��ÿ����Ŀ����
        For i = 0 To dsProjectSignature.Tables(0).Rows.Count - 1
            tmpProjectCode = dsProjectSignature.Tables(0).Rows(i).Item("project_code")

            '��ȡ����Ŀ��workflow_id
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")

            '�ж���Ŀ�Ƿ���ǩԼ
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature' and task_status='F'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            If dsTempAttend.Tables(0).Rows.Count > 0 Then
                '�׳�"���ܳ�����ǩԼ�ļƻ�"
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowRecordSignatureErr()
                Throw wfErr
                Exit Function
            End If

            '��������Ŀ�ĵǼ�ǩԼ�����ÿ�
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For j = 0 To dsTempAttend.Tables(0).Rows.Count - 1
                dsTempAttend.Tables(0).Rows(i).Item("task_status") = ""
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            '��Project-Task-Attendee��ȡ����Ŀ����һ�£�PlanSignature����״̬Ϊ��F��������
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='PlanSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")

            '������״̬��Ϊ��P����
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='PlanSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            dsTempAttend.Tables(0).Rows(0).Item("task_status") = "P"
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            '���Ǽ�ǩԼ��ʱ�����״̬��ΪDBNULL,��ʼʱ����ΪDBNULL
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'}"
            dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
            For j = 0 To dsTimingTask.Tables(0).Rows.Count - 1
                dsTimingTask.Tables(0).Rows(j).Item("status") = DBNull.Value
                'dsTimingTask.Tables(0).Rows(j).Item("start_time") = DBNull.Value
            Next
            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

        Next

        '��ȡ��Workflow-id����Ŀ����ƥ�����Ŀ����A�Ǻ�B��,�������Σ����ղ��������ﾭ��
        '�����Ŀ����A�ǲ��ڻ�������˼��У�����Ŀ����A�Ǽ����������˼���
        strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
        dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

        Dim ArrAttend As New ArrayList
        tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))
        ArrAttend.Add(tmpManagerA)

        '�����Ŀ����B�ǲ��ڻ�������˼��У�����Ŀ����B�Ǽ����������˼���
        tmpManagerB = Trim(dsTempAttend.Tables(0).Rows(0).Item("25"))
        ArrAttend.Add(tmpManagerB)
        tmpMinister = Trim(dsTempAttend.Tables(0).Rows(0).Item("31"))
        ArrAttend.Add(tmpMinister)
        tmpManagerLaw = Trim(dsTempAttend.Tables(0).Rows(0).Item("33"))

        '���ղ����ͷ��ﾭ��Ϊͬһ��
        If tmpManagerLaw <> tmpMinister Then
            ArrAttend.Add(tmpManagerLaw)
        End If

        tmpDirector = getTaskActor("01")
        ArrAttend.Add(tmpDirector)

        ''��	����ǩԼ�����˼������г�Ա
        ''����Ϣ����ӡ�Signature-plan-dateǩԼ��������Ϣ��
        'Dim newRow As DataRow
        'strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        'dsSignaturePlan = SignaturePlan.GetSignaturePlanInfo(strSql)

        ''�쳣����  
        'If dsSignaturePlan.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsSignaturePlan.Tables(0))
        '    Throw wfErr
        'End If

        'Dim tmpSignaturePlanDate As DateTime = dsSignaturePlan.Tables(0).Rows(0).Item("signature_plan_date")
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")
        'For i = 0 To ArrAttend.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = CStr(tmpSignaturePlanDate) & "ǩԼ����"
        '        .Item("accepter") = ArrAttend(i)
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)


        'ɾ���û����������Ŀ
        For i = 0 To dsProjectSignature.Tables(0).Rows.Count - 1
            dsProjectSignature.Tables(0).Rows(i).Item("signature_plan_code") = DBNull.Value
        Next
        ProjectSignature.UpdateProjectSignature(dsProjectSignature)

        '����ɾ��
        dsSignaturePlan.Tables(0).Rows(0).Delete()
        SignaturePlan.UpdateSignaturePlan(dsSignaturePlan)
    End Function


    '����ǩԼ�ƻ�����Ŀ
    Public Function CancelSignaturePlanProject(ByVal projectID As String)

        Dim strSql As String
        Dim dsTempAttend, dsTimingTask As DataSet
        Dim i As Integer

        '�ڲ����˱��н�����Ŀ�ļ�¼ǩԼ�ƻ�������Ϊ""
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordSignature'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '�ڲ����˱��н�����Ŀ�İ���ǩԼ������Ϊ"P"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='PlanSignature'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '���Ǽ�ǩԼ��ʱ�����״̬��ΪDBNULL
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordSignature'}"
        dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTimingTask.Tables(0).Rows.Count - 1
            dsTimingTask.Tables(0).Rows(i).Item("status") = DBNull.Value
            'dsTimingTask.Tables(0).Rows(i).Item("start_time") = DBNull.Value
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

    End Function

    '�����ϻ�
    Public Function ReMeetingPlan(ByVal projectID As String)
        Dim strSql As String
        Dim i As Integer
        Dim dsTemp As DataSet

        '2010-11-03 yjf add ��������ڼ�¼�������ۣ���Ŀ�������������ϻ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordReviewConclusion'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("task_status")), "", dsTemp.Tables(0).Rows(0).Item("task_status")) = "P" Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowMustRecord()
            Throw wfErr
        End If

        '����ǰ��������ر�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='P'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            '
            'qxd modify 2004-7-22 
            '
            '�����ϻ�ʱ����Ŀ��ǰ����״̬ΪP��������ΪA(��������״̬)��
            '�ڼ�¼�������۽������ӡ��ָ���Ŀ״̬��ѡ�ѡ��ָ���Ŀ״̬��
            '����Ŀ����״̬ΪA������״̬�û�P��������Ŀ��������ת��
            '���齫�Ƿ������ս���ѡ���Ϊ���ָ���Ŀ״̬����

            'dsTemp.Tables(0).Rows(i).Item("task_status") = ""
            dsTemp.Tables(0).Rows(i).Item("task_status") = "A"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '2009-4-20 yjf add
        '���ҵ��Ʒ��Ϊ���±���������Ŀ��������
        If dsTemp.Tables(0).Rows(0).Item("workflow_id") = "10" Then

            strSql = "{project_code='" & projectID & "' and task_id='ProjectAppraiseReport'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("task_status") = "P"
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("start_time") = Now
            Next
            WfProjectTask.UpdateWfProjectTask(dsTemp)

        Else

            '��ProjectProbe����ĺ�������ύ���н��ۣ���Ϊ"P"
            Dim strTaskID As String
            strTaskID = getProjectProbeNextTaskSQL(projectID, "ProjectProbe")
            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='SubmissionProbeResult'}" 'ApplyCapitialEvaluated,PreguaranteeActivity,ProjectAppraiseReport,ProjectAttitude,SubmissionProbeResult
            strSql = "{project_code=" & "'" & projectID & "'" & " and ( " & strTaskID & ") }"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("task_status") = "P"
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("start_time") = Now
            Next
            WfProjectTask.UpdateWfProjectTask(dsTemp)

            '2007-07-13 yjf add 
            '���������������״̬��Ϊ�գ����������ϻ�����õ������п�����Ϊ��״̬Ϊ��ɣ���������
            strSql = "{project_code='" & projectID & "' and task_id='ReviewMeetingPlan'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("task_status") = ""
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)
        End If

        '2005-10-26 yjf add 
        '����Ŀ�Ľ׶θ�Ϊ����״̬Ϊ�����ϻ�
        '����Ŀԭ��״̬��¼����Ŀ���У��Ա�ָ���Ŀ״̬ʱ�ָ���׶κ�״̬
        Dim ojbProject As New Project(conn, ts)
        strSql = "{project_code='" & projectID & "'}"
        dsTemp = ojbProject.GetProjectInfo(strSql)
        Dim tmpPhase, tmpStatus As String
        tmpPhase = dsTemp.Tables(0).Rows(0).Item("phase")
        tmpStatus = dsTemp.Tables(0).Rows(0).Item("status")

        dsTemp.Tables(0).Rows(0).Item("phase") = "����"
        dsTemp.Tables(0).Rows(0).Item("status") = "�����ϻ�"


        dsTemp.Tables(0).Rows(0).Item("origPhase") = tmpPhase
        dsTemp.Tables(0).Rows(0).Item("origStatus") = tmpStatus

        ojbProject.UpdateProject(dsTemp)

    End Function

    '���ĳ��Ŀ��ĳ������ĺ�������task_id

    Private Function getProjectProbeNextTaskSQL(ByVal projectID As String, ByVal taskID As String)
        Dim strSql, strTaskID As String
        Dim ds As DataSet
        Dim dr As DataRow
        Dim i, count As Integer

        strSql = "{project_code='" & projectID & "' and task_id='" & taskID & "'}"
        ds = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        strTaskID = "task_id='SubmissionProbeResult'"
        If Not ds Is Nothing Then
            count = ds.Tables(0).Rows.Count
            If count > 0 Then
                For i = 0 To count - 1
                    dr = ds.Tables(0).Rows(i)
                    With dr
                        If i = 0 Then
                            strTaskID = "task_id='" & .Item("next_task") & "' or "
                        ElseIf i = count - 1 Then
                            strTaskID = strTaskID & "task_id='" & .Item("next_task") & "'"
                        Else
                            strTaskID = strTaskID & "task_id='" & .Item("next_task") & "' or "
                        End If
                    End With
                Next
            End If
        End If
        Return strTaskID

    End Function

    '��γ���ſ�
    Public Function ReLoanApplication(ByVal projectID As String)
        Dim strSql As String
        Dim i As Integer
        Dim dsTemp As DataSet

        '������ſ���Ϊ"P"
        ''strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='LoanApplication'}"
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckedSignature'}"
        'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    dsTemp.Tables(0).Rows(i).Item("task_status") = "P"
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '2004-3-18 start
        'ͨ�������ύ���Ǽ�ǩԼ����RecordSignature��

        Dim strUser As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordSignature'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If dsTemp.Tables(0).Rows.Count > 0 Then
            strUser = dsTemp.Tables(0).Rows(0).Item("attend_person")
        End If

        finishedTask("", projectID, "RecordSignature", "", strUser)
        'end 
    End Function

    '��Ŀ���
    Public Function SplitPrjoect(ByVal fatherProjectID As String, ByVal sonProjectID As String)
        Dim strSql As String
        Dim i As Integer
        Dim newRow As DataRow
        Dim dsTempFather, dsTempSon As DataSet
        Dim dsAttend As DataSet
        Dim strWorkflowID As String

        ''��ȡ����Ŀ����Ʒ�ֵ�����ID
        'strSql = "{project_code='" & fatherProjectID & "' and task_id='RecordReviewConclusion'}"
        'dsTempFather = WfProjectTask.GetWfProjectTaskInfo(strSql)

        ''�쳣����  
        'If dsTempFather.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoRecordkErr(dsTempFather.Tables(0))
        '    Throw wfErr
        'End If

        '2011-9-1 yjf edit ��ȡ�����Ŀ������ID

        Dim dsTempProject As DataSet = project.GetProjectInfo("{project_code='" & sonProjectID & "'}")

        '�쳣����  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        strWorkflowID = dsTempProject.Tables(0).Rows(0).Item("split_workflow_id")


        Dim dsTemp, dsTemplate As DataSet

        Dim strWorkflow As String = "workflow_id=" & "'" & strWorkflowID & "'"

        '����ģ��
        dsTemplate = GetWfProjectTaskTemplateInfo("task_template", strWorkflow)
        dsTemp = WfProjectTask.GetWfProjectTaskInfo("null")

        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1


            newRow = dsTemp.Tables(0).NewRow()
            With newRow

                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("sequence") = dsTemplate.Tables(0).Rows(i).Item("sequence")
                .Item("task_name") = dsTemplate.Tables(0).Rows(i).Item("task_name")
                .Item("task_type") = dsTemplate.Tables(0).Rows(i).Item("task_type")
                .Item("apply_tool") = dsTemplate.Tables(0).Rows(i).Item("apply_tool")
                .Item("parameters") = dsTemplate.Tables(0).Rows(i).Item("parameters")
                .Item("duration") = dsTemplate.Tables(0).Rows(i).Item("duration")
                .Item("merge_relation") = dsTemplate.Tables(0).Rows(i).Item("merge_relation")
                .Item("flow_tool") = dsTemplate.Tables(0).Rows(i).Item("flow_tool")
                .Item("create_person") = dsTemplate.Tables(0).Rows(i).Item("create_person")
                .Item("create_date") = dsTemplate.Tables(0).Rows(i).Item("create_date")
                .Item("project_phase") = dsTemplate.Tables(0).Rows(i).Item("phase")
                .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                .Item("hasMessage") = dsTemplate.Tables(0).Rows(i).Item("hasMessage")

            End With

            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTask.UpdateWfProjectTask(dsTemp)

        '��ɫģ��
        dsTemplate = GetWfProjectTaskTemplateInfo("task_role_template", strWorkflow)
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo("null")

        '����ɫģ����ӵ���ɫ����
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '3��ת������ģ��
        dsTemplate = GetWfProjectTaskTemplateInfo("task_transfer_template", strWorkflow)
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")

        '��ת������ģ����ӵ�ת����������
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("next_task") = dsTemplate.Tables(0).Rows(i).Item("next_task")
                .Item("transfer_condition") = dsTemplate.Tables(0).Rows(i).Item("transfer_condition")
                .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                .Item("isItem") = dsTemplate.Tables(0).Rows(i).Item("isItem")
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)


        '4����ʱ����ģ��
        dsTemplate = GetWfProjectTaskTemplateInfo("timing_task_template", strWorkflow)
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")

        '������ģ����ӵ�����ģ��ʵ������
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("distance") = dsTemplate.Tables(0).Rows(i).Item("distance")
                .Item("start_time") = "1900-01-01"
                .Item("message_id") = dsTemplate.Tables(0).Rows(i).Item("message_id")
                .Item("type") = dsTemplate.Tables(0).Rows(i).Item("type")
                .Item("time_limit") = dsTemplate.Tables(0).Rows(i).Item("time_limit")
                .Item("parameter") = dsTemplate.Tables(0).Rows(i).Item("parameter")
                .Item("create_person") = dsTemplate.Tables(0).Rows(i).Item("create_person")
                .Item("create_date") = dsTemplate.Tables(0).Rows(i).Item("create_date")
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)


        '����Ŀ����A��B��Ϊ�յ�Ա����Ϊ��Ŀ����
        Dim tmpManagerA, tmpManagerB As String
        strSql = "select nowManagerA,nowManagerB from queryProjectInfo where projectCode='" & fatherProjectID & "'"
        dsTempFather = commQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsTempFather.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempFather.Tables(0))
            Throw wfErr
        End If

        tmpManagerA = dsTempFather.Tables(0).Rows(0).Item("nowManagerA")
        tmpManagerB = dsTempFather.Tables(0).Rows(0).Item("nowManagerB")


        strSql = "{project_code=" & "'" & sonProjectID & "'" & " and role_id='24' and attend_person=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpManagerA
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        strSql = "{project_code=" & "'" & sonProjectID & "'" & " and role_id='25' and attend_person=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpManagerB
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        '2011-9-1 yjf add ���ò����Ŀ��������¼Ա
        Dim strFatherAttendee, strFatherWorkflowID As String
        strSql = "{project_code=" & "'" & fatherProjectID & "'" & " and task_id='" & "RecordReviewConclusion" & "'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            Throw wfErr
        End If

        If Not dsTempTaskAttendee Is Nothing Then
            strFatherAttendee = dsTempTaskAttendee.Tables(0).Rows(0).Item("attend_person")
            strFatherWorkflowID = dsTempTaskAttendee.Tables(0).Rows(0).Item("workflow_id")
        End If

        Dim strMeetingRecordPerson As String
        strMeetingRecordPerson = strFatherAttendee
        'If strWorkflowID = "31" Or strWorkflowID = "32" Or strWorkflowID = "33" Or strWorkflowID = "34" Then
        '    strMeetingRecordPerson = "�·ﵤ"
        'Else
        '    If strFatherWorkflowID = "31" Or strFatherWorkflowID = "32" Or strFatherWorkflowID = "33" Or strFatherWorkflowID = "34" Then
        '        strMeetingRecordPerson = "��ʤ��"
        '    Else
        '        strMeetingRecordPerson = strFatherAttendee
        '    End If
        'End If

        strSql = "{project_code=" & "'" & sonProjectID & "'" & " and task_id='RecordReviewConclusion'}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim dsStaff As DataSet
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strMeetingRecordPerson
            dsAttend.Tables(0).Rows(i).Item("task_status") = "P"

            strSql = "{staff_name=" & "'" & strMeetingRecordPerson & "'" & "}"
            dsStaff = staff.FetchStaff(strSql)
            '�쳣����  
            If dsStaff.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsStaff.Tables(0))
                Throw wfErr
            End If
            dsStaff.Tables(0).Rows(0).Item("DoScan") = 1
            staff.UpdateStaff(dsStaff)
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)


        'qxd add 2004-7-23
        '������Ŀ�����˱�project_responsible

        dsTempSon = ProjectResponsible.GetProjectResponsibleInfo("null")

        strSql = "{project_code='" & fatherProjectID & "' }"
        dsTempFather = ProjectResponsible.GetProjectResponsibleInfo(strSql)
        For i = 0 To dsTempFather.Tables(0).Rows.Count - 1
            newRow = dsTempSon.Tables(0).NewRow()
            With newRow
                .Item("project_code") = sonProjectID
                .Item("manager_A") = dsTempFather.Tables(0).Rows(i).Item("manager_A")
                .Item("manager_B") = dsTempFather.Tables(0).Rows(i).Item("manager_B")
                .Item("create_person") = dsTempFather.Tables(0).Rows(i).Item("create_person")
                .Item("create_date") = dsTempFather.Tables(0).Rows(i).Item("create_date")
            End With
            dsTempSon.Tables(0).Rows.Add(newRow)
        Next
        ProjectResponsible.UpdateProjectResponsible(dsTempSon)

    End Function

    '������Ϣ
    Private Function AddMsg(ByVal projectID As String, ByVal taskID As String, ByVal msg As String, ByVal accepterID As String, ByVal respsonserID As String)

        '��	����Ϣ�������Ϣ����ʾ��Ϣ��Ա������N������
        '  �����Ϣ����
        Dim msgContent As String
        msgContent = respsonserID & " " & msg
        Dim dsTempTaskMessages As DataSet
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow
        ' �����Ϣ
        With newRow
            .Item("project_code") = projectID
            .Item("message_content") = msgContent
            .Item("accepter") = accepterID
            .Item("send_time") = Now
            .Item("is_affirmed") = "N"
        End With
        dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)

    End Function

    '��Ŀ��ͣʱ������Ϣ���ϼ�����:��������(01),��������(21),���ղ���(31)
    Private Sub sendMessageToManager(ByVal projectID As String, ByVal delayDay As Integer)
        Dim strSql As String
        Dim ds, dsCorporation As DataSet
        Dim i, count As Integer
        Dim strStaff, strCorporation As String

        strSql = "select distinct staff_name from staff_role where role_id in ('01','21','31')"
        ds = commQuery.GetCommonQueryInfo(strSql)

        strSql = "select corporation_name from corporation a " & _
                "left join  project b on a.corporation_code=b.corporation_code where b.project_code='" & projectID & "'"
        dsCorporation = commQuery.GetCommonQueryInfo(strSql)

        If Not dsCorporation Is Nothing Then
            If dsCorporation.Tables(0).Rows.Count > 0 Then
                strCorporation = dsCorporation.Tables(0).Rows(0).Item("corporation_name") & "��Ŀ(" & projectID & "):"
            Else
                strCorporation = "��Ŀ " & projectID
            End If
        End If

        If Not ds Is Nothing Then
            count = ds.Tables(0).Rows.Count
            If count > 0 Then
                For i = 0 To count - 1
                    strStaff = ds.Tables(0).Rows(i).Item("staff_name")
                    AddMsg(projectID, "", "��Ŀ��ͣ " & delayDay & " (��)", strStaff, strCorporation)
                Next
            End If
        End If
    End Sub

    '�ж�����״̬�Ƿ���ڡ�P��
    Private Function isTaskStatusEqualP(ByVal projectId As String, ByVal taskID As String) As Boolean
        Dim strSql As String
        Dim dsTemp As DataSet

        strSql = "{project_code='" & projectId & "' and task_id='" & taskID & "' and isnull(task_status,'')='P'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If dsTemp.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    '���ü�¼�������۵���ʾʱ��=����ʱ��+���ʱ��
    Private Sub setReviewConclusionCueTime(ByVal projectID As String, ByVal taskID As String, ByVal conferenceTime As DateTime)
        Dim strSql As String
        Dim i As Integer

        '�ڶ�ʱ������뵱ǰ����IDƥ��Ķ�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and type='A'}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)


        Dim tmpTimeLimit As Integer

        '����¼�������۶�ʱ����Ŀ�ʼʱ����Ϊ����Ļ���ʱ�䣫��ʾ���

        Dim newRow As DataRow
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            '��ʾ����
            tmpTimeLimit = dsTempTimingTask.Tables(0).Rows(i).Item("time_limit")
            newRow = dsTempTimingTask.Tables(0).Rows(i)
            With newRow
                .Item("start_time") = DateAdd(DateInterval.Hour, tmpTimeLimit * 24, conferenceTime) '����ʱ�䣽����ʱ�䣫��ʾ���
            End With
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)
    End Sub

    '��������
    Public Function updateProcess()

        '1�����isLiving=1����Ŀ���뼯�� 
        Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
        Dim ds As DataSet
        Dim strSql As String
        Dim count, i As Integer

        strSql = "select distinct project_code,workflow_id " & _
                 " from project_task " & _
                 " where project_code in (select project_code from project where isliving=1) order by project_code"
        ds = CommonQuery.GetCommonQueryInfo(strSql)
        If Not ds Is Nothing Then
            count = ds.Tables(0).Rows.Count
        Else
            Exit Function
        End If

        '2��
        Dim projectCode, workFlowID As String

        For i = 0 To count - 1
            projectCode = ds.Tables(0).Rows(i).Item("project_code")
            workFlowID = ds.Tables(0).Rows(i).Item("workflow_id")
            CommonQuery.PUpdateProcess(projectCode, workFlowID)
        Next
    End Function

    Public Function updateProcess(ByVal ProjectCode As String)

        '1�����isLiving=1����Ŀ���뼯�� 
        Dim CommonQuery As New BusinessRules.CommonQuery(conn, ts)
        Dim ds As DataSet
        Dim strSql As String
        Dim count, i As Integer

        strSql = "select distinct project_code,workflow_id " & _
                 " from project_task " & _
                 " where project_code in (select project_code from project where isliving=1 and project_code='" & ProjectCode & "') order by workflow_id"
        ds = CommonQuery.GetCommonQueryInfo(strSql)
        If Not ds Is Nothing Then
            count = ds.Tables(0).Rows.Count
        Else
            Exit Function
        End If

        '2��
        Dim workFlowID As String

        For i = 0 To count - 1
            'projectCode = ds.Tables(0).Rows(i).Item("project_code")
            workFlowID = ds.Tables(0).Rows(i).Item("workflow_id")
            CommonQuery.PUpdateProcess(ProjectCode, workFlowID)
        Next
    End Function
End Class
