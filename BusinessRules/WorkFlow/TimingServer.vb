Option Explicit On 

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'当工作流服务器每次启动或每天7：00时，调用方法TimingServer，检查启动定时任务。
Public Class TimingServer

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义任务列表引用
    Private WfProjectTask As WfProjectTask

    ''定义工作记录对象引用
    'Private WorkLog As WorkLog

    '定义项目对象引用
    Private Project As Project

    '定义消息字典对象引用
    Private WfMessagesTemplate As WfMessagesTemplate

    '定义工作流消息对象引用
    Private WfProjectMessages As WfProjectMessages

    '定义角色员工对象引用
    Private role As role

    '定义工作流对象引用
    Private WorkFlow As WorkFlow

    ''定义一个月的定时服务数组,布尔型变量
    'Private Ddone(30) As Boolean

    '定义假期对象引用
    Private Holiday As Holiday

    Private TimeServerLog As TimeServerLog

    Private Branch As Branch

    Private CommonQuery As CommonQuery

    '定义当天，本小时扫描的布尔变量
    Private bDay As Boolean
    Private bHour As Boolean

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction, ByVal isDayScan As Boolean, ByVal isHourScan As Boolean)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化定时任务对象
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        '实例化任务列表对象
        WfProjectTask = New WfProjectTask(conn, ts)

        '实例化项目对象
        Project = New Project(conn, ts)

        ''实例化工作记录对象
        'WorkLog = New WorkLog(conn, ts)

        '实例化消息字典对象
        WfMessagesTemplate = New WfMessagesTemplate(conn, ts)

        '实例化工作流消息对象
        WfProjectMessages = New WfProjectMessages(conn, ts)

        '实例化角色员工对象
        role = New Role(conn, ts)

        '实例化工作流消息对象
        WorkFlow = New WorkFlow(conn, ts)

        '实例化假期对象
        Holiday = New Holiday(conn, ts)

        TimeServerLog = New TimeServerLog(conn, ts)


        Branch = New Branch(conn, ts)

        CommonQuery = New CommonQuery(conn, ts)

        '将参数值赋予当天，本小时扫描的布尔变量
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

        '创建日志文件
        'Dim fs As FileStream = New FileStream("Log.txt", FileMode.OpenOrCreate) '打开并创建日志文件
        'Dim w As StreamWriter = New StreamWriter(fs) '创建写文件流
        'w.BaseStream.Seek(0, SeekOrigin.End) '指向文件末尾

        '①	获取系统日期；
        Dim sysTime As DateTime = Now

        '判断当天是否需要扫描（针对其他消息或任务）

        If bDay = False Then


            '获取早于等于Today的NoWorkingDay. Scaned=False的天数NoScanedDay；
            Dim sysDay As String = FormatDateTime(sysTime, DateFormat.ShortDate)
            strSql = "{holiday<=" & "'" & sysDay & "'" & " and isnull(scaned,0)<>1}"
            dsHoliday = Holiday.GetHolidayInfo(strSql)
            NoScanedDay = dsHoliday.Tables(0).Rows.Count

            '将早于Today的未扫描的NoWorkingDay的Scaned置为True;
            For i = 0 To dsHoliday.Tables(0).Rows.Count - 1

                dsHoliday.Tables(0).Rows(i).Item("scaned") = 1

            Next
            Holiday.UpdateHoliday(dsHoliday)

            '将定时任务表中任务类型为“A”，状态为“P”的定时任务开始时间改为开始时间+NoScanedDay×24;
            If NoScanedDay <> 0 Then
                strSql = "{type='A' and status='P'}"
                dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                    dsTemp.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Hour, NoScanedDay * 24, dsTemp.Tables(0).Rows(i).Item("start_time"))
                Next
                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)
            End If


            '②	获取定时任务状态为“P”,任务类型为“T,P,W”的任务；
            strSql = "{status='P' and type in ('T','P','W','R','M')}"
            dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

            '	对于每一个任务
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


                    '' 在任务角色获取与定时任务表PID、活动ID、提示角色匹配的员工和任务状态；
                    'strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
                    'dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    Select Case Trim(dsTempTimingTask.Tables(0).Rows(i).Item("type"))

                        Case "T" '        如果定时任务类型为("T")

                            '        如果（定时任务的开始日期=当前日期）
                            If tmpStartTime <= sysTime Then
                                '   在任务表获取与定时任务表PID、活动ID匹配的任务ID；
                                '   将定时任务状态置为“E”
                                '   调用startupTask(模板ID、PID、任务ID)启动转移任务；
                                '   如果任务表状态非空，将项目状态置为任务表的当前任务状态值；
                                '   如果任务表当前任务阶段非空，将项目阶段置为任务表的当前任务阶段值；


                                '2010-8-3 yjf add 根据在保余额是否为0作为是否逾期标准
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

                                        '2010-8-3 yjf add 发送逾期消息给所有的法务经理
                                        For Each drTemp As DataRow In dsTemp.Tables(0).Rows
                                            AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, drTemp.Item("staff_name"), tmpStaffID, "9", "N")
                                        Next

                                        '将项目的状态置为任务的相应状态
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


                                        '将定时任务状态置为“E”[用户只需提醒一次]；
                                        dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                    End If

                                End If

                            End If


                        Case "P" '        如果定时任务类型为("P")
                            '如果(定时任务的开始日期 = 当前日期)
                            '    将定时任务状态改为“E”；
                            '    在工作日志获取任务已完成次数HaveDone；
                            '    获取当前定时任务状态为“E”的任务次数ShouldDone；
                            '    获取该角色的任务员工；
                            '       对每个员工()
                            '         如果ShouldDone>HaveDone
                            '         AddMsg（项目ID，任务ID，消息ID、员工、“N”）；

                            If tmpStartTime <= sysTime Then

                                dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                '获取当前定时任务状态为“E”的任务次数ShouldDone；
                                strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id='" & tmpRoleID & "' and type='P' and status='E'}"
                                dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                                tmpPromptCount = dsTemp.Tables(0).Rows.Count

                                '统计已完成次数HaveDone记录数；
                                'strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                                'dsTempWorkLog = WorkLog.GetWorkLogInfo(strSql)
                                'tmpRecordCount = dsTempWorkLog.Tables(0).Rows.Count
                                strSql = "select COUNT(*) AS RecordCount from project_account_detail where  project_code = " & " '" & tmpProjectID & "'" & " and item_type='34' and item_code='001'"
                                tmpRecordCount = CommonQuery.GetCommonQueryInfo(strSql).Tables(0).Rows(0).Item("RecordCount")

                                ''获取该角色的任务员工；
                                ''       对每个员工()
                                ''         如果ShouldDone>HaveDone
                                ''         AddMsg（项目ID，任务ID，消息ID、员工、“N”）；
                                'For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                                '    tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                                '    If tmpPromptCount > tmpRecordCount Then
                                '        AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N")
                                '    End If
                                'Next

                                '2010-05-13 YJF ADD 
                                If tmpPromptCount > tmpRecordCount Then
                                    AddMessgeToRealAccepterForRefund(tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpRoleID, tmpMessageID, tmpStartTime)
                                    ''将定时任务状态置为“E”[用户只需提醒一次]；
                                    '2010-05-24 YJF EDIT 将定时任务状态置为“P”[如果未按时还款继续提示]；
                                    dsTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
                                    WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)
                                End If


                            End If

                        Case "W"
                                '自动恢复工作流任务处理
                                If tmpStartTime <= sysTime Then
                                    WorkFlow.resumeProcess(tmpProjectID)
                                End If

                                ''将定时任务状态置为“E”[用户只需提醒一次]；
                                'dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                                'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                        Case "R"
                                '循环启动的提示信息
                                '如果(定时任务的开始日期 = 当前日期)
                                '将该定时任务的开始时间置为任务的开始时间+任务的间隔时间
                                If tmpStartTime <= sysTime Then
                                    strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                                    dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                                    Dim tmpR As DateTime = IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("start_time")), Now, dsTempTimingTask.Tables(0).Rows(i).Item("start_time"))
                                    Dim tmpD As Integer = IIf(IsDBNull(dsTempTimingTask.Tables(0).Rows(i).Item("distance")), 0, dsTempTimingTask.Tables(0).Rows(i).Item("distance"))
                                    dsTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Day, tmpD, tmpR)
                                    WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                                End If

                                '2010-05-13 YJF ADD 任务责任是并非项目经理，但需要以项目经理作为责任人发送消息的情况
                        Case "M"
                                If tmpStartTime <= sysTime Then

                                    '2010-05-13 YJF ADD 
                                    AddMessgeToRealAccepterForManagerA(tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpRoleID, tmpMessageID)

                                    '2010-5-12 yjf edit 可循环提醒(如果任务未完成，间隔指定时间后继续提醒，否则，关闭定时提醒)
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


        '判断本小时是否需要扫描（针对超时消息）
        If bHour = False Then

            '②	获取定时任务状态为“P”,任务类型为“A”的任务；
            strSql = "{status='P' and type='A'}"
            dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

            '	对于每一个任务
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

                    '如果（定时任务的开始时间=当前时间）(对日期和小时同时判断)
                    If (tmpStartTime < sysTime) Or (tmpStartTime = sysTime And Hour(tmpStartTime) <= Hour(Now)) Then

                        '' 在任务角色获取与定时任务表PID、活动ID、提示角色匹配的员工和任务状态；
                        'strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
                        'dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


                        ''如果员工和任务状态集为空，获取提示角色的员工(领导)，向消息库添加消息（项目ID、任务ID，“逾期”，员工ID）；
                        'If dsTempTaskAttenddee.Tables(0).Rows.Count = 0 Then

                        '    'tmpStaffID = WorkFlow.getTaskActor(tmpRoleID)

                        '    '2010-05-11 yjf edit 按部门获取角色人员
                        '    strSql = "{project_code='" & tmpProjectID & "'}"
                        '    dsTempProject = Project.GetProjectInfo(strSql)
                        '    tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name"))
                        '    tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

                        '    If tmpStaffID = "" Then

                        '        '获取分支机构的上级机构

                        '        strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                        '        dsBranch = Branch.GetBranch(strSql)

                        '        tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                        '        '获取上级机构的参与人ACTOR
                        '        tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

                        '    End If

                        '    '获取任务的责任人
                        '    strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                        '    dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                        '    '向领导发送每个责任人的提示消息
                        '    If dsTempTaskAttenddee.Tables(0).Rows.Count <> 0 Then
                        '        For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                        '            tmpResponser = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                        '            AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N")
                        '        Next
                        '    Else
                        '        '向领导发送每个责任人的提示消息
                        '        AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, "", tmpMessageID, "N")
                        '    End If
                        'Else
                        '    '        否则(非空)
                        '    '            获取每该角色状态为“P”的员工；
                        '    '向消息库添加消息（项目ID、任务ID，“逾期”，员工ID）；

                        '    For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                        '        tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                        '        If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")) = "P" Then
                        '            AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N")
                        '        End If
                        '    Next
                        'End If


                        AddMessgeToRealAccepter(tmpProjectID, tmpWorkFlowID, tmpTaskID, tmpRoleID, tmpMessageID)

                        ''将定时任务状态置为“E”[用户只需提醒一次]；
                        'dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                        'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                        '2010-5-12 yjf edit 可循环提醒(如果任务未完成，间隔指定时间后继续提醒，否则，关闭定时提醒)
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

        '更新定时器日值
        TimeServerLog.UpdateTimeServerLog(dsTimeServerLog)

        ''关闭日志文件对象
        'w.Close()
        'fs.Close()

    End Function

    Public Function AddMsg(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal tmpStaffID As String, ByVal messageID As String, ByVal readFlag As String)
        Dim strSql, tmpTaskName, tmpMsgType, msgContent As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet
        '①	获取项目ID的名称；
        '②	获取任务ID的任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '③	获取消息ID的消息名称；
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '异常处理
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '获取项目的企业名称
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '④	如果消息ID的消息类型为“Project“ ，用项目ID+消息名称作为提示消息；
        Select Case tmpMsgType
            Case "project"
                msgContent = tmpCorporationName & "项目" & msgContent
                '⑤	如果消息ID的消息类型为“Task“ ，用项目ID+任务名称+消息名称作为提示消息；
            Case "task"
                msgContent = tmpCorporationName & "项目的" & tmpTaskName & "任务:" & msgContent
        End Select

        '⑥	在消息库添加消息（提示消息、员工、“N”）；
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

        '①	获取项目ID的名称；
        '②	获取任务ID的任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '③	获取消息ID的消息名称；
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '异常处理
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '获取项目的企业名称
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '④	如果消息ID的消息类型为“Project“ ，用项目ID+消息名称作为提示消息；
        Select Case tmpMsgType
            Case "project"
                msgContent = responserID & " " & tmpCorporationName & "项目" & msgContent
                '⑤	如果消息ID的消息类型为“Task“ ，用项目ID+任务名称+消息名称作为提示消息；
            Case "task"
                msgContent = responserID & " " & tmpCorporationName & "项目的" & tmpTaskName & "任务:" & msgContent
        End Select

        '⑥	在消息库添加消息（提示消息、员工、“N”）；
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
        '①	获取项目ID的名称；
        '②	获取任务ID的任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")


        '获取项目的企业名称
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        messageContent = tmpCorporationName & "项目" & messageContent

        '⑥	在消息库添加消息（提示消息、员工、“N”）；
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

        ' 在任务角色获取与定时任务表PID、活动ID、提示角色匹配的员工和任务状态；
        strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
        dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '2010-8-3 yjf add 如果为登记还款证明书任务,则发送超时消息给所有法务经理
        If tmpTaskID = "RecordRefundCertificate" And tmpRoleID = "33" Then
            strSql = "select staff_name from staff_role where isnull(overdue_message,0)=1"
            dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

            '2010-8-3 yjf add 发送逾期消息给所有的法务经理
            For Each drTemp As DataRow In dsTemp.Tables(0).Rows
                AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, drTemp.Item("staff_name"), tmpStaffID, tmpMessageID, "N")
            Next

            Exit Function
        End If


        '如果员工和任务状态集为空，获取提示角色的员工(领导)，向消息库添加消息（项目ID、任务ID，“逾期”，员工ID）；
        If dsTempTaskAttenddee.Tables(0).Rows.Count = 0 Then

            'tmpStaffID = WorkFlow.getTaskActor(tmpRoleID)

            '2010-05-11 yjf edit 按部门获取角色人员
            strSql = "select deptName from queryProjectInfo where ProjectCode='" & tmpProjectID & "'"
            'dsTempProject = Project.GetProjectInfo(strSql)
            dsTempProject = CommonQuery.GetCommonQueryInfo(strSql)
            tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("deptName")), "", dsTempProject.Tables(0).Rows(0).Item("deptName"))
            tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

            If tmpStaffID = "" Then

                '获取分支机构的上级机构

                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                dsBranch = Branch.GetBranch(strSql)

                tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '获取上级机构的参与人ACTOR
                tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

            End If

            '获取任务的责任人
            strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '向领导发送每个责任人的提示消息
            If dsTempTaskAttenddee.Tables(0).Rows.Count <> 0 Then
                For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                    tmpResponser = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                    AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N")
                Next
            Else
                '向领导发送每个责任人的提示消息
                AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, "", tmpMessageID, "N")
            End If
        Else
            '        否则(非空)
            '            获取每该角色状态为“P”的员工；
            '向消息库添加消息（项目ID、任务ID，“逾期”，员工ID）；

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

        '如果角色不为项目经理,即向领导发送消息
        If tmpRoleID <> "24" Then

            '2010-05-11 yjf edit 按部门获取角色人员
            strSql = "{project_code='" & tmpProjectID & "'}"
            dsTempProject = Project.GetProjectInfo(strSql)
            tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name"))
            tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

            If tmpStaffID = "" Then

                '获取分支机构的上级机构

                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                dsBranch = Branch.GetBranch(strSql)

                tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '获取上级机构的参与人ACTOR
                tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

            End If

            AddHeaderMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N")

        Else
            '否则向项目经理本人发送消息
            AddMsg(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpResponser, tmpMessageID, "N")

        End If

    End Function


    Public Function AddHeaderMsgForRefund(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal accepterID As String, ByVal responserID As String, ByVal messageID As String, ByVal readFlag As String, ByVal tmpStartTime As Date)
        Dim strSql, tmpTaskName, tmpMsgType, msgContent As String
        Dim dsTempTask, dsTempMsgTemplate, dsTempTaskMessages As DataSet

        '①	获取项目ID的名称；
        '②	获取任务ID的任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '③	获取消息ID的消息名称；
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '异常处理
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '获取项目的企业名称
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '④	如果消息ID的消息类型为“Project“ ，用项目ID+消息名称作为提示消息；
        Select Case tmpMsgType
            Case "project"
                msgContent = responserID & " " & tmpCorporationName & "项目" & " " & tmpStartTime.ToShortDateString & "应" & msgContent
                '⑤	如果消息ID的消息类型为“Task“ ，用项目ID+任务名称+消息名称作为提示消息；
            Case "task"
                msgContent = responserID & " " & tmpCorporationName & "项目的" & tmpTaskName & "任务:" & msgContent
        End Select

        '⑥	在消息库添加消息（提示消息、员工、“N”）；
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
        '①	获取项目ID的名称；
        '②	获取任务ID的任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
        '③	获取消息ID的消息名称；
        strSql = "{message_id=" & "'" & messageID & "'" & "}"
        dsTempMsgTemplate = WfMessagesTemplate.GetWfMessagesTemplateInfo(strSql)

        '异常处理
        If dsTempMsgTemplate.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempMsgTemplate.Tables(0))
            Throw wfErr
        End If

        tmpMsgType = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_type")
        msgContent = dsTempMsgTemplate.Tables(0).Rows(0).Item("message_content")

        '获取项目的企业名称
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)
        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")
        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '④	如果消息ID的消息类型为“Project“ ，用项目ID+消息名称作为提示消息；
        Select Case tmpMsgType
            Case "project"
                msgContent = tmpCorporationName & "项目" & " " & tmpStartTime.ToShortDateString & "应" & msgContent
                '⑤	如果消息ID的消息类型为“Task“ ，用项目ID+任务名称+消息名称作为提示消息；
            Case "task"
                msgContent = tmpCorporationName & "项目的" & tmpTaskName & "任务:" & msgContent
        End Select

        '⑥	在消息库添加消息（提示消息、员工、“N”）；
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

        ' 在任务角色获取与定时任务表PID、活动ID、提示角色匹配的员工和任务状态；
        strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & " and role_id=" & "'" & tmpRoleID & "'" & "}"
        dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        '如果员工和任务状态集为空，获取提示角色的员工(领导)，向消息库添加消息（项目ID、任务ID，“逾期”，员工ID）；
        If dsTempTaskAttenddee.Tables(0).Rows.Count = 0 Then

            'tmpStaffID = WorkFlow.getTaskActor(tmpRoleID)

            '2010-05-11 yjf edit 按部门获取角色人员
            strSql = "{project_code='" & tmpProjectID & "'}"
            dsTempProject = Project.GetProjectInfo(strSql)
            tmpBranch = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTempProject.Tables(0).Rows(0).Item("applicantTeam_name"))
            tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpBranch)

            If tmpStaffID = "" Then

                '获取分支机构的上级机构

                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                dsBranch = Branch.GetBranch(strSql)

                tmpSuper = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '获取上级机构的参与人ACTOR
                tmpStaffID = WorkFlow.getTaskActor(tmpProjectID, tmpTaskID, tmpRoleID, tmpSuper)

            End If

            '获取任务的责任人
            strSql = "{project_code=" & "'" & tmpProjectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTaskAttenddee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '向领导发送每个责任人的提示消息
            If dsTempTaskAttenddee.Tables(0).Rows.Count <> 0 Then
                For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                    tmpResponser = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                    AddHeaderMsgForRefund(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpResponser, tmpMessageID, "N", tmpStartTime)
                Next
            Else
                '向领导发送每个责任人的提示消息
                AddHeaderMsgForRefund(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, "", tmpMessageID, "N", tmpStartTime)
            End If
        Else
            '        否则(非空)
            '            获取每该角色状态为“P”的员工；
            '向消息库添加消息（项目ID、任务ID，“逾期”，员工ID）；

            For j = 0 To dsTempTaskAttenddee.Tables(0).Rows.Count - 1
                tmpStaffID = dsTempTaskAttenddee.Tables(0).Rows(j).Item("attend_person")
                If IIf(IsDBNull(dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")), "", dsTempTaskAttenddee.Tables(0).Rows(j).Item("task_status")) = "P" Then
                    AddMsgForRefund(tmpWorkFlowID, tmpProjectID, tmpTaskID, tmpStaffID, tmpMessageID, "N", tmpStartTime)
                End If
            Next
        End If

    End Function

End Class
