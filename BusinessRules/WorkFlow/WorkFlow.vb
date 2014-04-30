Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WorkFlow

    '定义模板常量
    Public Const Table_Task_Template As String = "task_template"
    Public Const Table_Task_Transfer_Template As String = "task_transfer_template"
    Public Const Table_Task_Role_Template As String = "task_role_template"
    Public Const Table_Timing_Task_Template As String = "timing_task_template"


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_WfTemplate As SqlDataAdapter

    '定义查询命令
    Private GetWfTemplateInfoCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '定义工作流对象引用
    Private WfProjectTask As WfProjectTask
    Private WfProjectMessages As WfProjectMessages
    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTimingTask As WfProjectTimingTask
    Private WfProjectTrack As WfProjectTrack
    Private ProjectResponsible As ProjectResponsible


    '定义项目引用
    Private project As project

    '定义工作日志对象引用
    Private WorkLog As WorkLog

    Private WorkflowType As WorkflowType

    Private commQuery As CommonQuery

    Private staff As staff

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        dsCommand_WfTemplate = New SqlDataAdapter()

        '引用外部事务
        ts = trans

        '实例化工作流对象
        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectMessages = New WfProjectMessages(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        WfProjectTrack = New WfProjectTrack(conn, ts)

        ProjectResponsible = New ProjectResponsible(conn, ts)

        '实例化项目对象
        project = New Project(conn, ts)

        '实例化工作日志对象
        WorkLog = New WorkLog(conn, ts)

        WorkflowType = New WorkflowType(conn, ts)

        commQuery = New CommonQuery(conn, ts)

        staff = New Staff(conn, ts)

    End Sub

    '获取工作流模板信息
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

    '向消息库中发送消息
    Private Function SendMessage()

    End Function

    '启动任务
    '2005-03-18 yjf add 设置前置任务，前置任务ID和前置任务责任人
    Public Function StartupTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal rollbackMsg As String, ByVal responserID As String, ByVal submitTaskID As String, ByVal submitUser As String)

        '调用原启动任务方法
        StartupTask(workFlowID, projectID, taskID, rollbackMsg, responserID)

        Dim strSql As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
            '2005-03-18 yjf add 设置前置任务，前置任务ID和前置任务责任人
            dsTempTaskStatus.Tables(0).Rows(i).Item("previous_task_id") = submitTaskID
            dsTempTaskStatus.Tables(0).Rows(i).Item("previous_task_attendee") = submitUser
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)
    End Function

    '启动任务
    Public Function StartupTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal rollbackMsg As String, ByVal responserID As String)

        '在任务表获取与参数(模板ID、项目ID、任务ID)匹配的任务
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '获取系统时间；在任务表任务的开始时间改为系统时间
        Dim sysTime As DateTime = Now

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        dsTempTask.Tables(0).Rows(0).Item("start_time") = sysTime

        '如果启动任务的启动模式为"Manual",将其置空
        Dim tmpStartMode As String = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_mode")), "", dsTempTask.Tables(0).Rows(0).Item("start_mode")))
        If tmpStartMode = "manual" Then
            dsTempTask.Tables(0).Rows(0).Item("start_mode") = DBNull.Value
        End If

        WfProjectTask.UpdateWfProjectTask(dsTempTask)


        '获取任务是否需发送消息
        Dim bTmp As Boolean = False
        If Not dsTempTask.Tables(0).Rows(0).Item("hasMessage") Is DBNull.Value Then
            bTmp = dsTempTask.Tables(0).Rows(0).Item("hasMessage")
        End If

        Dim tmpAttend, tmpBranch As String


        '如果任务参与人为空(非分配角色)(有委托权限的角色有且只有一条任务记录)


        '  调用getTaskActor（RoleID）获取任务参与人；
        '  将当前任务的参与人置为获取的任务参与人；

        Dim dsTempTaskAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理 
        If dsTempTaskAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttend.Tables(0))
            Throw wfErr
        End If

        'tmpAttend = IIf(IsDBNull(dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person")), "", dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person"))
        'Dim tmpRoleID As String = dsTempTaskAttend.Tables(0).Rows(0).Item("role_id")
        'Dim staff As New Staff(conn, ts)


        ''发送回退消息
        'If rollbackMsg <> "" Then
        '    AddMsg(projectID, taskID, rollbackMsg, tmpAttend, responserID)
        'End If

        'If tmpAttend = "" Then

        '    '获取起始任务的参与人的分支机构；
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and  task_type='BEGIN'}"
        '    Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '    '异常处理  
        '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '        Dim wfErr As New WorkFlowErr()
        '        wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '        Throw wfErr
        '    End If

        '    Dim tmpBeginTask As String = dsTemp.Tables(0).Rows(0).Item("task_id")
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpBeginTask & "'" & "}"
        '    dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    '异常处理  
        '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '        Dim wfErr As New WorkFlowErr()
        '        wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '        Throw wfErr
        '    End If

        '    tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
        '    strSql = "{staff_name=" & "'" & tmpAttend & "'" & "}"
        '    dsTemp = staff.FetchStaff(strSql)

        '    '异常处理  
        '    If dsTemp.Tables(0).Rows.Count = 0 Then
        '        Dim wfErr As New WorkFlowErr()
        '        wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '        Throw wfErr
        '    End If


        '    tmpBranch = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("branch_name")), "", dsTemp.Tables(0).Rows(0).Item("branch_name"))

        '    '调用getTaskActor（RoleID，分支机构）获取任务参与人ACTOR；
        '    tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpBranch)

        '    '如果ACTOR为空，调用getTaskActor（RoleID）获取任务参与人ACTOR；
        '    If tmpAttend = "" Then

        '        '获取分支机构的上级机构
        '        Dim Branch As New Branch(conn, ts)
        '        strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
        '        Dim dsBranch As DataSet = Branch.GetBranch(strSql)

        '        '异常处理  
        '        If dsBranch.Tables(0).Rows.Count = 0 Then
        '            Dim wfErr As New WorkFlowErr()
        '            wfErr.ThrowNoRecordkErr(dsBranch.Tables(0))
        '            Throw wfErr
        '        End If

        '        Dim tmpSuper As String = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

        '        '获取上级机构的参与人ACTOR
        '        tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpSuper)

        '        'tmpAttend = getTaskActor(tmpRoleID)
        '    End If

        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        '    dsTempTaskAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person") = tmpAttend
        '    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttend)
        'End If

        'qxd modify 2004-10-11
        '为了解决一个任务有两个角色做的问题。
        Dim k, count As Integer

        count = dsTempTaskAttend.Tables(0).Rows.Count

        Dim dsTemp As DataSet

        For k = 0 To count - 1
            tmpAttend = IIf(IsDBNull(dsTempTaskAttend.Tables(0).Rows(k).Item("attend_person")), "", dsTempTaskAttend.Tables(0).Rows(k).Item("attend_person"))
            Dim tmpRoleID As String = dsTempTaskAttend.Tables(0).Rows(k).Item("role_id")

            '发送回退消息
            If rollbackMsg <> "" Then
                AddMsg(projectID, taskID, rollbackMsg, tmpAttend, responserID)
            End If

            If tmpAttend = "" Then

                ''获取起始任务的参与人的分支机构；
                'strSql = "{project_code=" & "'" & projectID & "'" & " and  task_type='BEGIN'}"
                'Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

                ''异常处理  
                'If dsTemp.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                '    Throw wfErr
                'End If

                'Dim tmpBeginTask As String = dsTemp.Tables(0).Rows(0).Item("task_id")
                'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpBeginTask & "'" & "}"
                'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                ''异常处理  
                'If dsTemp.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                '    Throw wfErr
                'End If

                'tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
                'strSql = "{staff_name=" & "'" & tmpAttend & "'" & "}"
                'dsTemp = staff.FetchStaff(strSql)

                ''异常处理  
                'If dsTemp.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                '    Throw wfErr
                'End If

                strSql = "{project_code='" & projectID & "'}"
                dsTemp = project.GetProjectInfo(strSql)

                '异常处理  
                If dsTemp.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr()
                    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                    Throw wfErr
                End If

                tmpBranch = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("applicantTeam_name")), "", dsTemp.Tables(0).Rows(0).Item("applicantTeam_name"))

                '调用getTaskActor（RoleID，分支机构）获取任务参与人ACTOR；
                tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpBranch)

                '如果ACTOR为空，调用getTaskActor（RoleID）获取任务参与人ACTOR；
                If tmpAttend = "" Then

                    '获取分支机构的上级机构
                    Dim Branch As New Branch(conn, ts)
                    strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                    Dim dsBranch As DataSet = Branch.GetBranch(strSql)

                    '异常处理  
                    If dsBranch.Tables(0).Rows.Count = 0 Then
                        Dim wfErr As New WorkFlowErr()
                        wfErr.ThrowNoRecordkErr(dsBranch.Tables(0))
                        Throw wfErr
                    End If

                    Dim tmpSuper As String = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                    '获取上级机构的参与人ACTOR
                    tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpSuper)

                    'tmpAttend = getTaskActor(tmpRoleID)
                End If

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
                dsTempTaskAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                dsTempTaskAttend.Tables(0).Rows(k).Item("attend_person") = tmpAttend
                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttend)


            End If

            '2007-10-16 yjf edit 更新消息与任务刷新标记位
            If tmpAttend <> "" Then

                strSql = "{staff_name='" & tmpAttend & "'}"
                dsTemp = staff.FetchStaff(strSql)

                '异常处理  
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


        '将每个参与人的转移任务状态置为“P”
        Dim TimingServer As New TimingServer(conn, ts, True, True)
        Dim tmpUserID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
            tmpUserID = Trim(dsTempTaskStatus.Tables(0).Rows(i).Item("attend_person"))
            dsTempTaskStatus.Tables(0).Rows(i).Item("task_status") = "P"
            '发送消息
            If bTmp Then
                TimingServer.AddMsg(workFlowID, projectID, taskID, tmpUserID, "16", "N")
            End If
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)

        '在定时活动查找与当前任务ID匹配的定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and type='A'}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)


        Dim tmpTimeLimit As Integer

        '将所有定时任务的开始时间置为任务的开始时间+提示时间
        '将所有定时任务表的状态置为“P”；
        Dim newRow As DataRow
        '2010-8-3 yjf add 逾期消息及登记还款证明书消息除外（因为这个两个逾期消息在ImplAddRefundPlan接口添加进去以后，会启动任务，启动任务的时候会将其预警消息的启动时间推迟，置为：启动时间＝当前时间＋提示间隔）
        If taskID <> "OverdueRecord" And taskID <> "RecordRefundCertificate" Then

            For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                '提示期限
                tmpTimeLimit = dsTempTimingTask.Tables(0).Rows(i).Item("time_limit")
                newRow = dsTempTimingTask.Tables(0).Rows(i)
                With newRow
                    .Item("start_time") = DateAdd(DateInterval.Hour, tmpTimeLimit * 24, sysTime) '启动时间＝当前时间＋提示间隔
                    .Item("status") = "P"
                End With
            Next

        End If

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


    End Function

    '启动任务
    Public Function StartupManualTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal rollbackMsg As String, ByVal responserID As String)

        '在任务表获取与参数(模板ID、项目ID、任务ID)匹配的任务
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '获取系统时间；在任务表任务的开始时间改为系统时间
        Dim sysTime As DateTime = Now

        '异常处理
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        dsTempTask.Tables(0).Rows(0).Item("start_time") = sysTime
        WfProjectTask.UpdateWfProjectTask(dsTempTask)


        '获取任务是否需发送消息
        Dim bTmp As Boolean = False
        If Not dsTempTask.Tables(0).Rows(0).Item("hasMessage") Is DBNull.Value Then
            bTmp = dsTempTask.Tables(0).Rows(0).Item("hasMessage")
        End If

        Dim tmpAttend, tmpBranch As String


        '如果任务参与人为空(非分配角色)(有委托权限的角色有且只有一条任务记录)


        '  调用getTaskActor（RoleID）获取任务参与人；
        '  将当前任务的参与人置为获取的任务参与人；

        Dim dsTempTaskAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理 
        If dsTempTaskAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttend.Tables(0))
            Throw wfErr
        End If

        tmpAttend = IIf(IsDBNull(dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person")), "", dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person"))
        Dim tmpRoleID As String = dsTempTaskAttend.Tables(0).Rows(0).Item("role_id")
        Dim staff As New Staff(conn, ts)


        '发送回退消息
        If rollbackMsg <> "" Then
            AddMsg(projectID, taskID, rollbackMsg, tmpAttend, responserID)
        End If

        If tmpAttend = "" Then

            '获取起始任务的参与人的分支机构；
            strSql = "{project_code=" & "'" & projectID & "'" & " and  task_type='BEGIN'}"
            Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            Dim tmpBeginTask As String = dsTemp.Tables(0).Rows(0).Item("task_id")
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpBeginTask & "'" & "}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
            strSql = "{staff_name=" & "'" & tmpAttend & "'" & "}"
            dsTemp = staff.FetchStaff(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If


            tmpBranch = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("branch_name")), "", dsTemp.Tables(0).Rows(0).Item("branch_name"))

            '调用getTaskActor（RoleID，分支机构）获取任务参与人ACTOR；
            tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpBranch)

            '如果ACTOR为空，调用getTaskActor（RoleID）获取任务参与人ACTOR；
            If tmpAttend = "" Then

                '获取分支机构的上级机构
                Dim Branch As New Branch(conn, ts)
                strSql = "{branch_name=" & "'" & tmpBranch & "'" & "}"
                Dim dsBranch As DataSet = Branch.GetBranch(strSql)

                '异常处理  
                If dsBranch.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr()
                    wfErr.ThrowNoRecordkErr(dsBranch.Tables(0))
                    Throw wfErr
                End If

                Dim tmpSuper As String = IIf(IsDBNull(dsBranch.Tables(0).Rows(0).Item("superior_branch")), "", dsBranch.Tables(0).Rows(0).Item("superior_branch"))

                '获取上级机构的参与人ACTOR
                tmpAttend = getTaskActor(projectID, taskID, tmpRoleID, tmpSuper)

                'tmpAttend = getTaskActor(tmpRoleID)
            End If

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
            dsTempTaskAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            dsTempTaskAttend.Tables(0).Rows(0).Item("attend_person") = tmpAttend
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttend)
        End If


        '将每个参与人的转移任务状态置为“P”
        Dim TimingServer As New TimingServer(conn, ts, True, True)
        Dim tmpUserID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTaskStatus As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTempTaskStatus.Tables(0).Rows.Count - 1
            tmpUserID = Trim(dsTempTaskStatus.Tables(0).Rows(i).Item("attend_person"))
            dsTempTaskStatus.Tables(0).Rows(i).Item("task_status") = "P"
            '发送消息
            If bTmp Then
                TimingServer.AddMsg(workFlowID, projectID, taskID, tmpUserID, "16", "N")
            End If
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskStatus)

        '在定时活动查找与当前任务ID匹配的定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and type='A'}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)


        Dim tmpTimeLimit As Integer

        '将所有定时任务的开始时间置为任务的开始时间+提示时间
        '将所有定时任务表的状态置为“P”；
        Dim newRow As DataRow
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            '提示期限
            tmpTimeLimit = dsTempTimingTask.Tables(0).Rows(i).Item("time_limit")
            newRow = dsTempTimingTask.Tables(0).Rows(i)
            With newRow
                .Item("start_time") = DateAdd(DateInterval.Hour, tmpTimeLimit * 24, sysTime) '启动时间＝当前时间＋提示间隔
                .Item("status") = "P"
            End With
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


    End Function

    '创建工作流
    Public Function CreateProcess(ByVal workFlowID As String, ByVal projectID As String, ByVal userID As String) As String
        CreateProcess(workFlowID, projectID, userID, "1")
    End Function

    '创建工作流
    Public Function CreateProcess(ByVal workFlowID As String, ByVal projectID As String, ByVal userID As String, ByVal phase As String) As String


        '获取系统时间
        Dim sysTime As DateTime = Today

        Dim strSql As String

        ''获取项目阶段
        Dim tmpTaskPhase, tmpTaskStatus As String
        Dim dsTempProject As DataSet
        If phase = "1" Then
            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)
            '异常处理  
            If dsTempProject.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                Throw wfErr
            End If

            tmpTaskPhase = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("phase")), "", dsTempProject.Tables(0).Rows(0).Item("phase"))
        Else
            tmpTaskPhase = phase
        End If


        '根据业务品种和项目阶段获取模版ID
        strSql = "{service_type=" & "'" & workFlowID & "'" & " and isnull(phase,'')=" & "'" & tmpTaskPhase & "'" & "}"
        Dim dsWorkflowType As DataSet = WorkflowType.GetWorkflowTypeInfo(strSql)

        '异常处理  
        If dsWorkflowType.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsWorkflowType.Tables(0))
            Throw wfErr
        End If

        Dim strWorkflowID As String = dsWorkflowType.Tables(0).Rows(0).Item("workflow_id")
        Dim strWorkflow As String = "workflow_id=" & "'" & strWorkflowID & "'"
        'Dim strWorkflow As String = "workflow_id='01'"

        '在任务表查找是否存在方法参数中指定的工作流对象
        '如果不存在，创建工作流对象，否则异常处理；
        strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id=" & "'" & strWorkflowID & "'" & "}"
        Dim dsTemp As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)
        If dsTemp.Tables(0).Rows.Count <> 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowExistWorkFlowErr()
            Throw wfErr
        Else

            '1、任务模板
            Dim dsTemplate As DataSet = GetWfProjectTaskTemplateInfo("task_template", strWorkflow)

            Dim newRow As DataRow
            Dim i As Integer
            Dim straTime As DateTime = Now
            Dim beginTaskID As String

            '将担保业务工作流任务模板任务对象添加到任务表
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1


                newRow = dsTemp.Tables(0).NewRow()
                With newRow

                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '将所有添加任务的工作流ID置为项目编码
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

                '获取工作流的起始活动任务
                If IIf(IsDBNull(dsTemplate.Tables(0).Rows(i).Item("task_type")), "", dsTemplate.Tables(0).Rows(i).Item("task_type")) = "BEGIN" Then
                    beginTaskID = Trim(dsTemplate.Tables(0).Rows(i).Item("task_id"))
                    newRow.Item("start_time") = Now
                End If
                dsTemp.Tables(0).Rows.Add(newRow)
            Next

            WfProjectTask.UpdateWfProjectTask(dsTemp)

            '2、角色模板
            dsTemplate = GetWfProjectTaskTemplateInfo("task_role_template", strWorkflow)
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo("null")

            '将角色模板添加到角色表中
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
                newRow = dsTemp.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '将所有添加任务的工作流ID置为项目编码
                    .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                    .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")

                    '将工作流起始活动任务的员工ID属性置为咨询人员ID,并把它的任务状态置为“P”（进行）
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

            '3、转移条件模版
            dsTemplate = GetWfProjectTaskTemplateInfo("task_transfer_template", strWorkflow)
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")

            '将转移条件模版添加到转移条件表中
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
                newRow = dsTemp.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '将所有添加任务的工作流ID置为项目编码
                    .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                    .Item("next_task") = dsTemplate.Tables(0).Rows(i).Item("next_task")
                    .Item("transfer_condition") = dsTemplate.Tables(0).Rows(i).Item("transfer_condition")
                    .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                    .Item("isItem") = dsTemplate.Tables(0).Rows(i).Item("isItem")
                End With
                dsTemp.Tables(0).Rows.Add(newRow)
            Next

            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)


            '4、定时任务模板
            dsTemplate = GetWfProjectTaskTemplateInfo("timing_task_template", strWorkflow)
            dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")

            '将任务模板添加到任务模板实例表中
            For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
                newRow = dsTemp.Tables(0).NewRow()
                With newRow
                    .Item("workflow_id") = strWorkflowID
                    .Item("project_code") = projectID    '将所有添加任务的工作流ID置为项目编码
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


            '	如果当前任务的项目阶段非空，将项目阶段置为当前任务的阶段值；
            Dim dsTempTask As DataSet

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & beginTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '异常处理  
            If dsTempTask.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                Throw wfErr
            End If

            tmpTaskPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))
            tmpTaskStatus = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_status")), "", dsTempTask.Tables(0).Rows(0).Item("project_status")))


            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)

            '异常处理  
            If dsTempProject.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                Throw wfErr
            End If

            If tmpTaskPhase <> "" Then
                dsTempProject.Tables(0).Rows(0).Item("phase") = tmpTaskPhase
            End If

            '	如果当前任务的项目状态非空，将项目状态置为当前任务的状态

            If tmpTaskStatus <> "" Then
                dsTempProject.Tables(0).Rows(0).Item("status") = tmpTaskStatus
            End If

            project.UpdateProject(dsTempProject)


            finishedTask(strWorkflowID, projectID, beginTaskID, ".T.", userID)



        End If

    End Function

    '复制签约数据
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

    '复制放款回执数据
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

    '复制解保数据
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

    '结束任务
    Public Function finishedTask(ByVal workFlowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal finishedFlag As String, ByVal userID As String)
        'Dim strSql As String
        'Dim strMatchProjectCode As String

        'strSql = "select isnull(match_project_code,'') as match_project_code from project where project_code='" + projectID + "'"
        'Dim objCommonQuery As New CommonQuery(conn, ts)
        'Dim dsTemp As DataSet = objCommonQuery.GetCommonQueryInfo(strSql)
        'If Trim(dsTemp.Tables(0).Rows(0).Item("match_project_code")) <> "" And finishedTaskID <> "RecordReviewConclusion" Then
        '    strMatchProjectCode = Trim(dsTemp.Tables(0).Rows(0).Item("match_project_code"))

        '    '对应的贷款担保项目需签发放款通知书
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

        '    '对应的贷款担保项目需收取保费
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

        '    ''对应的贷款担保项目需收取保费
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

        '    '对应的小贷项目需登记放款回执
        '    If finishedTaskID = "RecordReturnReceipt" Then

        '        If workFlowID = "02" Then


        '            strSql = "select task_id from work_log where project_code='" + strMatchProjectCode + "' and task_id='ValidateLoanSmall'"
        '            dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)

        '            '对应的小贷项目需签发放款通知书
        '            If dsTemp.Tables(0).Rows.Count = 0 Then
        '                Dim wfErr As New WorkFlowErr
        '                wfErr.ThrowMustValidateLoanSmall()
        '                Throw wfErr
        '                Exit Function
        '            End If

        '            strSql = "select workflow_id from project_task_attendee where project_code='" + strMatchProjectCode + "' and task_id='ValidateLoanSmall'"
        '            dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)

        '            '小贷额度不需登记放款回执，所以不检查
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

        '        '复制放款信息
        '        CopyReturnReceipt(projectID, strMatchProjectCode)

        '    End If


        '    '复制放款记录
        '    If finishedTaskID = "LoanPetition" Then
        '        CopySignature(projectID, strMatchProjectCode)
        '    End If

        '    '对应的贷款担保项目需制作合同
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

        '    '复制签约信息
        '    If finishedTaskID = "RecordSignature" Then
        '        CopySignature(projectID, strMatchProjectCode)
        '    End If

        '    '复制解保信息
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
        '    '小贷额度项下需额度生效、有剩余额度以及额度未到期才可受理
        '    If workFlowID = "34" Then

        '        '如果credit_project_code为空则为受理申请，需增加设置credit_project_code为有效小贷额度的编码
        '        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        '        Dim dsTempProject As DataSet = project.GetProjectInfo(strSql)
        '        If dsTempProject.Tables(0).Rows(0).Item("credit_project_code") Is DBNull.Value Then
        '            '获取当前有效的小贷额度项目
        '            strSql = "select projectcode ,RemnantCredit from SmallCreditInfo where substring(projectcode,1,5)='" & projectID.Substring(0, 5) & "' order by applydate desc"
        '            dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '            If dsTemp.Tables(0).Rows.Count <> 0 Then
        '                If dsTemp.Tables(0).Rows(0).Item("RemnantCredit") > dsTempProject.Tables(0).Rows(0).Item("apply_sum") Then
        '                    dsTempProject.Tables(0).Rows(0).Item("credit_project_code") = dsTemp.Tables(0).Rows(0).Item("projectcode")
        '                    project.UpdateProject(dsTempProject)
        '                End If
        '            Else
        '                '无有效小贷授信额度或额度不足
        '                Dim wfErr As New WorkFlowErr
        '                wfErr.ThrowNoSmallCredit()
        '                Throw wfErr
        '                Exit Function
        '            End If
        '        End If
        '    End If

        '    '关联贷前小贷项目需还款后才可放款
        '    If workFlowID = "02" Then
        '        strSql = "select ServiceType,relate_project_code from queryProjectInfo where project_code='" & projectID & "'"
        '        dsTemp = objCommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp.Tables(0).Rows(0).Item("ServcieType") = "贷前小贷" Then
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

        '如果该任务已提交过,

        '1、如果当前任务ID不存在，抛出工作任务布存在异常
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            'Dim wfErr As New WorkFlowErr
            'wfErr.ThrowNotExistTaskErr()
            'Throw wfErr
            Exit Function
        Else


            '获取当前任务的名称,任务类型和开始时间
            tmpTaskName = dsTempTask.Tables(0).Rows(0).Item("task_name")
            tmpTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))
            startTime = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_time")), Now, dsTempTask.Tables(0).Rows(0).Item("start_time"))
            tmpTaskPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))
            tmpTaskStatus = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_status")), "", dsTempTask.Tables(0).Rows(0).Item("project_status")))
            tmpStartMode = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("start_mode")), "", dsTempTask.Tables(0).Rows(0).Item("start_mode")))

            '读取项目阶段和项目状态
            Dim tmpProjectPhase, tmpProjectStatus As String
            Dim dsTempProject As DataSet
            'strSql = "{project_code=" & "'" & projectID & "'" & "}"
            'dsTempProject = project.GetProjectInfo(strSql)

            ''异常处理  
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

            '2、如果当前任务ID状态为“W”，抛出不能提交暂停的任务异常
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
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

            '3、如果任务类型=“ISLAND”，取系统完成时间作为当前任务的完成时间，返回；
            If tmpTaskType = "ISLAND" Then

                '获取工作日志中该项目该任务并且AUTO类型为0的日志
                Dim Worklog As New WorkLog(conn, ts)
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and auto=0}"
                Dim dsWorklog As DataSet = Worklog.GetWorkLogInfo(strSql)

                For i = 0 To dsWorklog.Tables(0).Rows.Count - 1

                    newRow = dsWorklog.Tables(0).Rows(i)
                    With newRow

                        '将任务状态置为F,AUTO置为1,项目阶段和状态置为相应值
                        .Item("task_status") = "F"
                        .Item("project_phase") = tmpProjectPhase

                        '如果任务表中有状态，则把该状态记录到工作日志中，否则把项目的状态记录到工作日志中
                        If tmpTaskStatus = "" Then

                            .Item("project_status") = tmpProjectStatus

                            ''添加工作日志
                            'AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpProjectStatus, tmpStartMode)

                        Else
                            .Item("project_status") = tmpTaskStatus

                            'AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpTaskStatus, tmpStartMode)
                        End If

                    End With
                Next

                Worklog.UpdateWorkLog(dsWorklog)

                ''如果TaskID的任务参与者均完成此任务()
                'Dim isDone As Boolean = True
                'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                '    If Trim(IIf(IsDBNull(dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status"))) <> "F" Then
                '        isDone = False
                '    End If
                'Next

                ''   将满足Project_Code= ProjectID 、StartupTask= TaskID、Status=“P”条件的Project_Track对象的Status置为“F”；
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


            '获取系统时间
            Dim sysTime As DateTime = Now

            '2008-8-20 yjf add 增加工作流功能:同一任务如果有多人操作,则只要其中一人提交该任务即完成
            If tmpTaskType = "OPT" Then

                '7、[有效的完成了任务]在任务表将当前任务（ProjectID、TaskID、EmployeeID）状态改为“F”完成
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'}"
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '异常处理  
                If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
                    Throw wfErr
                End If

                Dim tempOptPerson As String
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    tempOptPerson = dsTempTaskAttendee.Tables(0).Rows(i).Item("attend_person")
                    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
                    '调用AACKMassage（ProjectID、TaskID、EmployeeID）自动确认任务消息；
                    AACKMassage(workFlowID, projectID, finishedTaskID, tempOptPerson)
                    '将员工完成任务时间置为系统时间
                    dsTempTaskAttendee.Tables(0).Rows(0).Item("end_time") = sysTime
                Next


            Else
                '7、[有效的完成了任务]在任务表将当前任务（ProjectID、TaskID、EmployeeID）状态改为“F”完成
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and attend_person=" & "'" & userID & "'" & "}"
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '异常处理  
                If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
                    Throw wfErr
                End If

                dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status") = "F"

                '8,	调用AACKMassage（ProjectID、TaskID、EmployeeID）自动确认任务消息；
                AACKMassage(workFlowID, projectID, finishedTaskID, userID)

                '将员工完成任务时间置为系统时间
                dsTempTaskAttendee.Tables(0).Rows(0).Item("end_time") = sysTime

            End If

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee) '更新操作

            '9、添加工作日志(当任务类型不为结束是,添加状态为空的工作日志,待任务流转后填充状态)
            'If isTaskStatusEqualP(projectID, finishedTaskID) Then
            If tmpTaskType <> "END" Then
                AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, "", tmpStartMode)
            End If
            'End If

            '10、如果（ProjectID、TaskID）的所有任务状态均为“F” 将定时任务中的当前任务ID的状态置为“E”；
            '否则，返回；

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            If flag = 0 Then
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '判断是否所有任务状态均为“F”
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    If dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") <> "F" Then
                        Exit Function
                    End If
                Next
            End If
            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            ''判断是否所有任务状态均为“F”
            'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            '    If dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") <> "F" Then
            '        Exit Function
            '    End If
            'Next

            '将定时任务中的当前任务ID的状态置为“E”
            Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
            If dsTempTimingTask.Tables(0).Rows.Count <> 0 Then
                For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                    dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
                Next
            End If

            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

            '[提交手动控制任务]如果当前任务的start_mode=“manual”，将start_mode置空，返回
            If tmpStartMode = "manual" Then
                dsTempTask.Tables(0).Rows(0).Item("start_mode") = ""
                WfProjectTask.UpdateWfProjectTask(dsTempTask)
                Exit Function
            End If

            '2005-04-27 yjf 修改：手动任务提交不改变项目状态
            '将项目的阶段置为任务的相应阶段
            strSql = "{project_code=" & "'" & projectID & "'" & "}"
            dsTempProject = project.GetProjectInfo(strSql)

            '异常处理  
            If dsTempProject.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                Throw wfErr
            End If

            If tmpTaskPhase <> "" Then
                dsTempProject.Tables(0).Rows(0).Item("phase") = tmpTaskPhase
            End If

            project.UpdateProject(dsTempProject)

            '11、如果当前任务提供流程工具，调用流程工具
            Dim tmpFlowTools As String
            Dim args As Object() = {conn, ts}
            If Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("flow_tool")), "", dsTempTask.Tables(0).Rows(0).Item("flow_tool"))) <> "" Then
                tmpFlowTools = dsTempTask.Tables(0).Rows(0).Item("flow_tool")
                tmpFlowTools = "BusinessRules." & tmpFlowTools
                tmpFlowTools = tmpFlowTools.Trim

                '动态创建接口对象
                Dim t As System.Type = System.Type.GetType(tmpFlowTools)
                Dim iFlowTools As IFlowTools = Activator.CreateInstance(t, args)
                iFlowTools.UseFlowTools(workFlowID, projectID, finishedTaskID, finishedFlag, userID)

            End If

            '4、获取当前任务的转移任务和转移条件记录集
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & "}"
            Dim dsTempTaskTransfer As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '5、 获取转移条件为真的转移任务集
            Dim nextTaskID, tmpTransCondition, tmpTransPhase, tmpTransStatus As String
            Dim dsConditionTrue As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")
            For i = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
                nextTaskID = Trim(dsTempTaskTransfer.Tables(0).Rows(i).Item("next_task"))
                tmpTransCondition = Trim(dsTempTaskTransfer.Tables(0).Rows(i).Item("transfer_condition"))
                tmpTransStatus = Trim(IIf(IsDBNull(dsTempTaskTransfer.Tables(0).Rows(i).Item("project_status")), "", dsTempTaskTransfer.Tables(0).Rows(i).Item("project_status")))

                '判断条件是否为真
                If CompareExpression(workFlowID, projectID, finishedTaskID, finishedFlag, tmpTransCondition) Then

                    '构造转移条件为真的转移任务集
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

            '6、如果转移条件为真的转移任务集为空
            If dsConditionTrue.Tables(0).Rows.Count = 0 Then

                '如果当前活动类型不是结束任务，抛出提交任务的结果无效错误,返回。
                If tmpTaskType <> "END" Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowInvalidSubmit()
                    Throw wfErr
                    Exit Function
                End If
            End If



            '12、如果当前任务的类型为“结束”，返回；
            If tmpTaskType = "END" Then

                '如果任务表中有状态，则把该状态记录到工作日志中，否则把项目的状态记录到工作日志中
                ' If isTaskStatusEqualP(projectID, finishedTaskID) Then
                If tmpTaskStatus = "" Then
                    '添加工作日志
                    AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpProjectStatus, tmpStartMode)
                Else
                    AddWorkLog(projectID, finishedTaskID, tmpTaskName, userID, "F", startTime, Now, 1, tmpProjectPhase, tmpTaskStatus, tmpStartMode)
                End If
                'End If

                Exit Function

            End If

            '13、[真正完成了一个有后继活动的任务]对每个转移条件为真的转移任务（多个任务）

            '   如果汇集类型为“AND”，获取转移任务的前趋活动状态；

            Dim tmpNextTaskID As String
            Dim tmpPreTaskID As String
            Dim dsTempPreTaskStatus, dsTempWorkLog As DataSet
            Dim tmpApplyTool As String

            '用FiliterJeeDummyTask（ProjectID、ShiftTaskSet）获取转移条件为真的业务任务（实任务）
            dsConditionTrue = FiliterJeeDummyTask(workFlowID, projectID, dsConditionTrue, finishedFlag, userID)

            '如果实任务集记录为空,证明为项目结束,将工作日志中该任务状态改为项目状态
            If dsConditionTrue.Tables(0).Rows.Count = 0 Then

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & finishedTaskID & "'" & " and project_status=''}"
                dsTempWorkLog = Worklog.GetWorkLogInfo(strSql)

                '重新获取项目状态
                strSql = "{project_code=" & "'" & projectID & "'" & "}"
                dsTempProject = project.GetProjectInfo(strSql)
                tmpProjectStatus = Trim(IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("status")), "", dsTempProject.Tables(0).Rows(0).Item("status")))

                For j = 0 To dsTempWorkLog.Tables(0).Rows.Count - 1
                    dsTempWorkLog.Tables(0).Rows(j).Item("project_status") = tmpProjectStatus
                Next
                Worklog.UpdateWorkLog(dsTempWorkLog)

            End If

            '遍历所有转移条件为真的转移任务集
            Dim tmpNextTaskType, mergeRelation As String
            Dim dsPreTaskMode As DataSet
            Dim tmpPreTaskMode As String

            '2005-09-13 yjf add 修改在有多个后续任务时,由于第一个后续任务前置任务未完成而导致第二个后续任务不处理的情况
            Dim isFinishedPreTask As Boolean

            For i = 0 To dsConditionTrue.Tables(0).Rows.Count - 1

                isFinishedPreTask = True

                tmpNextTaskID = dsConditionTrue.Tables(0).Rows(i).Item("next_task")
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpNextTaskID & "'" & "}"
                dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                '异常处理  
                If dsTempTask.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                    Throw wfErr
                End If


                tmpNextTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))
                mergeRelation = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("merge_relation")), "", dsTempTask.Tables(0).Rows(0).Item("merge_relation"))

                tmpTransStatus = Trim(IIf(IsDBNull(dsConditionTrue.Tables(0).Rows(i).Item("project_status")), "", dsConditionTrue.Tables(0).Rows(i).Item("project_status")))

                '调用AddTaskTrackRecord（ProjectID,Workflow_id,TaskID,StartupTask）记录流程历程信息;
                AddTaskTrackRecord("", projectID, finishedTaskID, tmpNextTaskID)

                '如果转移任务的项目状态属性值非空
                '   将项目的状态修改为转移任务的项目状态属性值；
                strSql = "{project_code=" & "'" & projectID & "'" & "}"
                dsTempProject = project.GetProjectInfo(strSql)

                '异常处理  
                If dsTempProject.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                    Throw wfErr
                End If

                If tmpTransStatus <> "" Then
                    dsTempProject.Tables(0).Rows(0).Item("status") = tmpTransStatus
                    project.UpdateProject(dsTempProject)
                End If

                '获取项目的状态；
                '在工作日志将TaskID任务的项目状态为空的记录的项目状态置为项目的状态值;
                dsTempProject = project.GetProjectInfo(strSql)

                '异常处理  
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

                '获取每一个转移条件为真的转移任务ID

                If mergeRelation = "AND" Then
                    '如果汇集类型为“AND”，获取转移任务的前趋活动状态；

                    strSql = "{project_code=" & "'" & projectID & "'" & " and next_task=" & "'" & tmpNextTaskID & "'" & "}"
                    '获取转移任务的前趋活动集
                    dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    '遍历转移任务的前趋活动集
                    For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1

                        tmpPreTaskID = dsTempTaskTransfer.Tables(0).Rows(j).Item("task_id")
                        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpPreTaskID & "'" & "}"

                        '获取每个前趋活动任务角色及其完成状态集
                        dsTempPreTaskStatus = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                        '遍历前趋活动的完成状态（是否有为P并且启动模式不为Manual）
                        For k = 0 To dsTempPreTaskStatus.Tables(0).Rows.Count - 1

                            '获取前趋任务的启动模式
                            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpPreTaskID & "'" & "}"
                            dsPreTaskMode = WfProjectTask.GetWfProjectTaskInfo(strSql)

                            '异常处理  
                            If dsPreTaskMode.Tables(0).Rows.Count = 0 Then
                                Dim wfErr As New WorkFlowErr
                                wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
                                Throw wfErr
                            End If

                            tmpPreTaskMode = IIf(IsDBNull(dsPreTaskMode.Tables(0).Rows(0).Item("start_mode")), "", dsPreTaskMode.Tables(0).Rows(0).Item("start_mode"))

                            If IIf(IsDBNull(dsTempPreTaskStatus.Tables(0).Rows(k).Item("task_status")), "", dsTempPreTaskStatus.Tables(0).Rows(k).Item("task_status")) <> "F" And tmpPreTaskMode <> "manual" Then
                                '2005-09-13 yjf add 修改在有多个后续任务时,由于第一个后续任务前置任务未完成而导致第二个后续任务不处理的情况
                                isFinishedPreTask = False
                                'Exit Function
                            End If
                        Next
                    Next

                    '2005-09-13 yjf add 修改在有多个后续任务时,由于第一个后续任务前置任务未完成而导致第二个后续任务不处理的情况
                    If isFinishedPreTask Then

                        '启动转移任务
                        StartupTask(workFlowID, projectID, tmpNextTaskID, "", "", finishedTaskID, userID)

                    End If

                Else
                    If mergeRelation = "XOR" Then
                        '启动转移任务后返回
                        StartupTask(workFlowID, projectID, tmpNextTaskID, "", "", finishedTaskID, userID)
                        Exit Function
                    Else
                        '启动转移任务
                        StartupTask(workFlowID, projectID, tmpNextTaskID, "", "", finishedTaskID, userID)
                    End If
                End If

                'If tmpNextTaskType = "AUTO" Then

                '    '获取应用工具
                '    strSql = "{project_code=" & "'" & projectID & "'"   & " and task_id=" & "'" & tmpNextTaskID & "'" & "}"
                '    dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
                '    tmpApplyTool = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("apply_tool")), "", dsTempTask.Tables(0).Rows(0).Item("apply_tool")))

                '    If tmpApplyTool <> "" Then
                '        '调用工具
                '        tmpApplyTool = "BusinessRules." & tmpApplyTool
                '        Dim t As System.Type = System.Type.GetType(tmpApplyTool)
                '        Dim iApplyTools As IApplyTools = Activator.CreateInstance(t, args)
                '        iApplyTools.UseApplyTools()

                '    End If

                '    '当工作流的虚活动执行完成工具后，调用此方法继续流转
                '    VTask(workFlowID, projectID, tmpNextTaskID, userID)

                'End If
            Next

        End If

    End Function

    '判断工作流是否挂起
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

    '流程挂起
    Public Function suspendProcess(ByVal projectID As String, ByVal delayDay As Integer)

        '①	将指定工作流任务状态为“P”的任务状态置为“W”；
        Dim i, j As Integer
        Dim sysTime As DateTime = Now
        Dim strSql As String
        Dim dsTempAttendee, dsTempTask, dstTempTimingTask As DataSet
        Dim tmpTaskID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='P'" & "}"
        dsTempAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttendee.Tables(0).Rows.Count - 1

            dsTempAttendee.Tables(0).Rows(i).Item("task_status") = "W"

            '②	任务的暂停开始时间置为系统时间；
            tmpTaskID = Trim(dsTempAttendee.Tables(0).Rows(i).Item("task_id"))
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
            For j = 0 To dsTempTask.Tables(0).Rows.Count - 1
                dsTempTask.Tables(0).Rows(j).Item("pause_start_time") = sysTime
            Next
            WfProjectTask.UpdateWfProjectTask(dsTempTask)

        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttendee)


        '③	获取定时任务表中指定工作流定时任务的定时类型为A，状态为“P”的定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and type='A' and status='P'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        '④	将定时任务状态置为“W”
        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "W"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)

        '将定时任务表中的恢复工作流任务的状态置为'P',开始时间置为当前时间+间隔时间
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='WakeProject'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "P"
            dstTempTimingTask.Tables(0).Rows(i).Item("start_time") = DateAdd(DateInterval.Day, delayDay, Now)
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)

        '项目暂停:向上级主管发送消息
        sendMessageToManager(projectID, delayDay)


    End Function

    '流程恢复
    Public Function resumeProcess(ByVal projectID As String)
        '①	在任务表获取项目编码指定的任务状态为“W”的任务；
        Dim i, j As Integer
        Dim sysTime As DateTime = Now
        Dim strSql As String
        Dim dsTempAttendee, dsTempTask, dstTempTimingTask As DataSet
        Dim tmpTaskID As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='W'" & "}"
        dsTempAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '②	将状态为“W”的任务状态置为“P”；
        For i = 0 To dsTempAttendee.Tables(0).Rows.Count - 1
            dsTempAttendee.Tables(0).Rows(i).Item("task_status") = "P"

            '③	获取系统时间；
            '④	暂停结束时间置为系统时间；
            tmpTaskID = Trim(dsTempAttendee.Tables(0).Rows(i).Item("task_id"))
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
            For j = 0 To dsTempTask.Tables(0).Rows.Count - 1
                dsTempTask.Tables(0).Rows(j).Item("pause_end_time") = sysTime
            Next
            WfProjectTask.UpdateWfProjectTask(dsTempTask)
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttendee)

        '⑤	在定时任务表中将指定工作流定时任务状态为“W”的任务开始时间改为开始时间+(暂停结束时间-暂停开始时间)； 
        strSql = "{project_code=" & "'" & projectID & "'" & " and status='W'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        Dim tmpWorkFlowID As String
        Dim tmpStartTime, tmpPauseEndTime, tmpPauseStartTime As DateTime

        '⑥	将指定工作流定时任务状态为“W”的任务状态改为“P”； 
        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            tmpWorkFlowID = dstTempTimingTask.Tables(0).Rows(i).Item("workflow_id")
            tmpTaskID = dstTempTimingTask.Tables(0).Rows(i).Item("task_id")
            strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id=" & "'" & tmpWorkFlowID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '异常处理  
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

        '将定时任务表中的恢复工作流任务的状态置为'E'
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='WakeProject'" & "}"
        dstTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For i = 0 To dstTempTimingTask.Tables(0).Rows.Count - 1
            dstTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dstTempTimingTask)


    End Function

    '任务回退
    Public Function rollbackTask(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal userID As String, ByVal rollbackMsg As String)

        Dim i, j As Integer
        Dim strSql As String
        Dim dsTask, dsProjectTrack, dsWorkLog, dsProject, dsAttend, dsTimingTask As DataSet
        Dim tmpStartMode, tmpRollBackTask, tmpFinshedTask, tmpWorklogPhase, tmpWorklogStatus, tmpStartupTask, tmpTaskStatus As String
        Dim iRangUp As Integer

        '①	如果Project_Task.Start_Mode=“manual”,提示“不能回退手工启动的任务！”，返回；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理  
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

        '②	获取Workflow_id=工作流ID、StartupTask= TaskID、Status=“P”的Project_Track对象（Serial_Num最小者）；
        strSql = "{project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & taskID & "'" & " and isnull(status,'')='P'}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        '异常处理  
        If dsProjectTrack.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectTrack.Tables(0))
            Throw wfErr
        End If

        '③	获取Project_Track对象的FinishedTask属性（RollBackTask）；
        tmpRollBackTask = Trim(dsProjectTrack.Tables(0).Rows(0).Item("FinishedTask"))

        '④	获取Project_Track对象的Serial_Num属性(RangeUp);
        iRangUp = dsProjectTrack.Tables(0).Rows(0).Item("serial_num")

        '⑥	获取回退任务启动任务的所在位置（Serial_Num）
        strSql = "{project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & tmpRollBackTask & "'" & " and serial_num<" & "'" & iRangUp & "'" & " order by serial_num desc}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        '异常处理  
        If dsProjectTrack.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectTrack.Tables(0))
            Throw wfErr
        End If

        iRangUp = dsProjectTrack.Tables(0).Rows(0).Item("serial_num")
        tmpFinshedTask = dsProjectTrack.Tables(0).Rows(0).Item("FinishedTask")

        '⑦ 将Project_Track对象的Status属性置为“P”；
        dsProjectTrack.Tables(0).Rows(0).Item("Status") = "P"
        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

        '在工作日志获取Project_Track. FinishedTask的项目阶段和项目状态;
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpFinshedTask & "'" & " order by finish_time desc}"
        dsWorkLog = WorkLog.GetWorkLogInfo(strSql)

        '异常处理  
        If dsWorkLog.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsWorkLog.Tables(0))
            Throw wfErr
        End If

        tmpWorklogPhase = IIf(IsDBNull(dsWorkLog.Tables(0).Rows(0).Item("project_phase")), "", dsWorkLog.Tables(0).Rows(0).Item("project_phase"))
        tmpWorklogStatus = IIf(IsDBNull(dsWorkLog.Tables(0).Rows(0).Item("project_status")), "", dsWorkLog.Tables(0).Rows(0).Item("project_status"))

        '将项目状态和阶段更新为Project_Track. FinishedTask的项目阶段和项目状态
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsProject = project.GetProjectInfo(strSql)

        '异常处理  
        If dsProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProject.Tables(0))
            Throw wfErr
        End If

        dsProject.Tables(0).Rows(0).Item("phase") = tmpWorklogPhase
        dsProject.Tables(0).Rows(0).Item("status") = tmpWorklogStatus
        project.UpdateProject(dsProject)


        '⑧	将RollBackTask任务加入回退任务集RollBackSet；
        Dim ArrRollBackTask As New ArrayList
        ArrRollBackTask.Add(tmpRollBackTask)


        '⑨	对于SERIAL-NUM> RangeUp的每个Project_Track对象
        strSql = "{project_code=" & "'" & projectID & "'" & " and serial_num>" & "'" & iRangUp & "'" & " order by serial_num}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1

            tmpFinshedTask = dsProjectTrack.Tables(0).Rows(i).Item("FinishedTask")
            For j = 0 To ArrRollBackTask.Count - 1
                '如果Project_Track. FinishedTask IN RollBackSet
                If tmpFinshedTask = ArrRollBackTask.Item(j) Then
                    '将Project_Track. StartupTask任务加入回退任务集RollBackSet;
                    tmpStartupTask = dsProjectTrack.Tables(0).Rows(i).Item("StartupTask")
                    ArrRollBackTask.Add(tmpStartupTask)

                    '删除Project_Track对象
                    dsProjectTrack.Tables(0).Rows(i).Delete()

                    Exit For

                End If
            Next

        Next

        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

        '对于回退任务集RollBackSet中的每个任务RollBackSet(i)
        For i = 0 To ArrRollBackTask.Count - 1

            '如果RollBackSet(i) = rollbackTask
            If ArrRollBackTask(i) = tmpRollBackTask Then
                '   调用startupTask(模板ID、项目ID、RollBackTask)启动回退任务;
                StartupTask(workFlowID, projectID, ArrRollBackTask(i), rollbackMsg, userID)

            Else
                '否则()

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & ArrRollBackTask(i) & "'" & "}"
                dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                '      将任务状态置为“”;

                For j = 0 To dsAttend.Tables(0).Rows.Count - 1
                    dsAttend.Tables(0).Rows(j).Item("task_status") = ""
                Next
                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

                dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

                '      将任务的定时任务状态置为”E”;
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

    '终止流程
    Public Function cancelProcess(ByVal projectID As String)

        '①	删除任务参与人表中项目编码的任务；
        Dim i As Integer
        Dim dsTempTask, dsTempAttend, dsTempTimingTask, dsTempTrans, dsTempProject, dsWorkLog, dsProjectTrack As DataSet
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & "}"
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

        '②	删除定时任务中指定项目编码的定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


        '②	删除工作日志中的任务；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsWorkLog = WorkLog.GetWorkLogInfo(strSql)
        For i = 0 To dsWorkLog.Tables(0).Rows.Count - 1
            dsWorkLog.Tables(0).Rows(i).Delete()
        Next
        WorkLog.UpdateWorkLog(dsWorkLog)


        '删除任务跟踪表中的记录
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)
        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1
            dsProjectTrack.Tables(0).Rows(i).Delete()
        Next
        WfProjectTrack.UpdateWfProjectTrack(dsTempTimingTask)

        '③ 删除任务表中的任务；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)
        For i = 0 To dsTempTask.Tables(0).Rows.Count - 1

            dsTempTask.Tables(0).Rows(i).Delete()
        Next

        WfProjectTask.UpdateWfProjectTask(dsTempTask)


        '设置项目结束标识isliving=0
        dsTempProject = project.GetProjectInfo(strSql)

        '异常处理  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        dsTempProject.Tables(0).Rows(0).Item("isliving") = 0

        '设置项目的状态为项目阶段+"暂缓"
        dsTempProject.Tables(0).Rows(0).Item("status") = dsTempProject.Tables(0).Rows(0).Item("phase") & "暂缓"
        project.UpdateProject(dsTempProject)


    End Function


    '手工启动任务（WorkflowID、ProjectID、TaskID）
    Public Function StartTaskByManual(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String)
        Dim strSql As String
        Dim dsTempTask As DataSet
        '①	在任务表获取参数指定的任务，将start_mode置为“manual”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理  
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        dsTempTask.Tables(0).Rows(0).Item("start_mode") = "manual"
        WfProjectTask.UpdateWfProjectTask(dsTempTask)
        '②	调用StartupManualTask(WorkflowID、ProjectID、TaskID、“”)；
        StartupManualTask(workflowID, projectID, taskID, "", "")
    End Function

    '获取流程任务（WorkflowID、ProjectID）
    Public Function GetAllBusinessTasks(ByVal workflowID As String, ByVal projectID As String) As DataSet
        Dim strSql As String
        Dim dsTempAttend As DataSet
        '①	在任务角色表获取与参数WorkflowID、ProjectID匹配的所有任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        '②	返回任务名称列表；
        Return dsTempAttend
    End Function


    '删除工作流
    Public Function deleteProcess(ByVal projectID As String)

    End Function

    '更改流程
    Public Function modifiyProcess(ByVal projectID As String)

    End Function

    '查询消息信息
    Public Function LookUpMessage(ByVal strCondition_ProjectMessage As String) As DataSet
        'Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and accepter=" & "'" & userID & "'" & "}"
        Dim dsTempProjectMessage As DataSet = WfProjectMessages.GetWfProjectMessagesInfo(strCondition_ProjectMessage)
        Return dsTempProjectMessage
    End Function

    '查询进行中的任务
    Public Function LookUpWorking(ByVal projectID As String, ByVal userID As String)

    End Function

    '查询进行中的任务
    Public Function LookUpWorking(ByVal userID As String) As DataSet

        ''搜索状态为“P”进行中的项目，取出其项目ID和任务ID
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

        '''搜索指定项目和任务的任务列表
        ''For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        ''    projectID = Trim(dsTemp.Tables(0).Rows(i).Item("project_code"))
        ''    taskID = dsTemp.Tables(0).Rows(i).Item("task_id")
        ''    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id =" & "'" & taskID & "'" & " }"
        ''    dsTask.Merge(WfProjectTask.GetWfProjectTaskInfo(strSql))
        ''Next

        'Dim dsTask As DataSet = Me.commQuery.GetCommonQueryInfo(strSql)

        ''返回任务列表
        'Return dsTask

        Dim dsTask As DataSet = Me.commQuery.LookUpWorking(userID)
        Return dsTask

    End Function

    '查询进行中的任务
    Public Function LookUpWorkingEx(ByVal sql_Condition As String) As DataSet

        '搜索状态为“P”进行中的项目，取出其项目ID和任务ID
        Dim strSql As String = sql_Condition
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim dsTask As DataSet = WfProjectTask.GetWfProjectTaskInfo("null")
        Dim projectID, taskID As String

        Dim i As Integer

        '搜索指定项目和任务的任务列表
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            projectID = Trim(dsTemp.Tables(0).Rows(i).Item("project_code"))
            taskID = dsTemp.Tables(0).Rows(i).Item("task_id")
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id =" & "'" & taskID & "'" & "}"
            dsTask.Merge(WfProjectTask.GetWfProjectTaskInfo(strSql))
        Next

        '返回任务列表
        Return dsTask

    End Function

    '查询流程状态
    Public Function LookUpStatus(ByVal projectID As String)

    End Function

    '比较判断表达式是否为真
    Private Function CompareExpression(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean
        '定义条件处理接口
        Dim iCondition As ICondition

        'Select Case taskID

        '    Case "ReviewFeeCharge", "CashlossReview"   '评审费收入是否等于支出
        '        iCondition = New ImplIncomePayout(conn, ts)
        '        'Case "CheckApplyTimes" '申请次数>3
        '        '    iCondition = New ImplCommon()
        '        '    Return iCondition.GetResult(workFlowID, projectID, taskID, ".T.", transCondition)
        '        'Case "ValidateReviewConclusion" '评审费是否不足
        '        '    iCondition = New ImplTrialFee(conn, ts)
        '        ''Case "GuaranteeCharge" '担保费收入是否等于支出
        '        ''    iCondition = New ImplGuaranteeFee(conn, ts)
        '        'Case "RefundRecord"  '还款是否结束
        '        '    iCondition = New ImplEndReturn(conn, ts)
        '    Case Else  '一般的情况
        '        iCondition = New ImplCommon()

        'End Select

        '一般的情况
        iCondition = New ImplCommon

        Return iCondition.GetResult(workFlowID, projectID, taskID, expFlag, transCondition)

    End Function

    '调用工具
    Public Function VTask(ByVal workFlowID As String, ByVal ProjectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String)
        '1、在任务表获取与参数(模板ID、项目ID、任务ID)匹配的任务
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '异常处理  
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
            Throw wfErr
        End If

        Dim tmpTaskType As String = dsTempTask.Tables(0).Rows(0).Item("task_type")

        '2、如果当前任务提供流程工具，调用流程工具；
        Dim tmpFlowTools As String
        Dim args As Object() = {conn, ts}
        If IsDBNull(dsTempTask.Tables(0).Rows(0).Item("flow_tool")) = False Then
            tmpFlowTools = dsTempTask.Tables(0).Rows(0).Item("flow_tool")
            tmpFlowTools = "BusinessRules." & tmpFlowTools

            '动态创建接口对象
            Dim t As System.Type = System.Type.GetType(tmpFlowTools)
            Dim iFlowTools As IFlowTools = Activator.CreateInstance(t, args)
            iFlowTools.UseFlowTools(workFlowID, ProjectID, taskID, finishedFlag, userID)

        End If


        '3、获取转移任务集
        Dim dsTempTaskTransfer As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '4、 获取转移条件为真的转移任务集
        Dim i, j, k As Integer
        Dim newRow As DataRow
        Dim nextTaskID, tmpTransCondition As String
        Dim dsConditionTrue As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")
        For i = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
            nextTaskID = dsTempTaskTransfer.Tables(0).Rows(i).Item("next_task")
            tmpTransCondition = Trim(dsTempTaskTransfer.Tables(0).Rows(i).Item("transfer_condition"))

            '判断条件是否为真
            If CompareExpression(workFlowID, ProjectID, taskID, ".T.", tmpTransCondition) Then

                '构造转移条件为真的转移任务集
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

        '5、如果当前任务的类型为“结束”，返回；
        If tmpTaskType = "END" Then
            Exit Function
        End If

        '6、[真正完成了一个有后继活动的任务]对每个转移条件为真的转移任务（多个任务）：
        Dim dsTempPreTask, dsTempPreTaskStatus As DataSet
        Dim mergeRelation, tmpPreTaskID, tmpNextTaskID, tmpTransStatus As String
        For i = 0 To dsConditionTrue.Tables(0).Rows.Count - 1
            tmpNextTaskID = dsConditionTrue.Tables(0).Rows(i).Item("next_task")
            strSql = "{project_code=" & "'" & ProjectID & "'" & " and task_id=" & "'" & tmpNextTaskID & "'" & "}"
            dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

            '异常处理  
            If dsTempTask.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                Throw wfErr
            End If

            mergeRelation = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("merge_relation")), "", dsTempTask.Tables(0).Rows(0).Item("merge_relation"))

            tmpTransStatus = IIf(IsDBNull(dsConditionTrue.Tables(0).Rows(i).Item("project_status")), "", dsConditionTrue.Tables(0).Rows(i).Item("project_status"))

            If mergeRelation = "AND" Then
                '如果汇集类型为“AND”，获取转移任务的前趋活动状态；

                strSql = "{project_code=" & "'" & ProjectID & "'" & " and next_task=" & "'" & tmpNextTaskID & "'" & "}"
                '获取转移任务的前趋活动集
                dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                '遍历转移任务的前趋活动集
                For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1

                    tmpPreTaskID = dsTempTaskTransfer.Tables(0).Rows(j).Item("task_id")
                    strSql = "{project_code=" & "'" & ProjectID & "'" & " and task_id=" & "'" & tmpPreTaskID & "'" & "}"
                    '获取每个前趋活动任务角色及其完成状态集
                    dsTempPreTaskStatus = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    '遍历前趋活动的完成状态（是否全为“F”）
                    For k = 0 To dsTempPreTaskStatus.Tables(0).Rows.Count - 1
                        If dsTempPreTaskStatus.Tables(0).Rows(k).Item("task_status") <> "F" Then
                            Exit Function
                        End If
                    Next
                Next

                '启动转移任务
                StartupTask(workFlowID, ProjectID, tmpNextTaskID, "", "")

            Else

                '启动转移任务
                StartupTask(workFlowID, ProjectID, tmpNextTaskID, "", "")
            End If

        Next

    End Function

    '添加工作日志
    Public Function AddWorkLog(ByVal projectID As String, ByVal taskID As String, ByVal taskName As String, ByVal userID As String, ByVal taskStatus As String, ByVal startTime As DateTime, ByVal finishTime As DateTime, ByVal autoType As Integer, ByVal projectPhase As String, ByVal projectStatus As String, ByVal start_mode As String)
        Dim workLog As New WorkLog(conn, ts)
        Dim dsTempWorkLog As DataSet = workLog.GetWorkLogInfo("null")

        '获取该任务的角色ID
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsAttendRole As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
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

    'FiliterJeeDummyTask（ProjectID、ShiftTaskSet）
    '过滤后继活动集的虚活动，获取业务活动。
    Private Function FiliterJeeDummyTask(ByVal workflowID As String, ByVal projectID As String, ByVal ShiftTaskSet As DataSet, ByVal finishedFlag As String, ByVal userID As String) As DataSet

        '①	Vtask=True
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


                '②	对于转移条件为真的转移活动集ShiftTaskSet中的每个任务TaskID
                '    在任务表获取当前任务（ProjectID、TaskID）；

                tmpTaskID = Trim(ShiftTaskSet.Tables(0).Rows(i).Item("next_task"))
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"


                '获取任务的活动类型；
                dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                '异常处理  
                If dsTempTask.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                    Throw wfErr
                End If

                tmpTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))
                tmpPhase = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("project_phase")), "", dsTempTask.Tables(0).Rows(0).Item("project_phase")))


                '如果任务类型为“AUTO”
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

                    '如果应用工具非空，调用应用工具；
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                    dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                    '异常处理  
                    If dsTempTask.Tables(0).Rows.Count = 0 Then
                        Dim wfErr As New WorkFlowErr
                        wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                        Throw wfErr
                    End If

                    'tmpApplyTool = Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("apply_tool")), "", dsTempTask.Tables(0).Rows(0).Item("apply_tool")))

                    'If tmpApplyTool <> "" Then
                    '    '调用工具
                    '    tmpApplyTool = "BusinessRules." & tmpApplyTool
                    '    t = System.Type.GetType(tmpApplyTool)
                    '    Dim iApplyTools As IApplyTools = Activator.CreateInstance(t, args)
                    '    iApplyTools.UseApplyTools()

                    'End If

                    '如果当前任务提供流程工具，调用流程工具；
                    If Trim(IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("flow_tool")), "", dsTempTask.Tables(0).Rows(0).Item("flow_tool"))) <> "" Then

                        tmpFlowTools = dsTempTask.Tables(0).Rows(0).Item("flow_tool")
                        tmpFlowTools = "BusinessRules." & tmpFlowTools

                        '动态创建接口对象
                        t = System.Type.GetType(tmpFlowTools)
                        Dim iFlowTools As IFlowTools = Activator.CreateInstance(t, args)
                        iFlowTools.UseFlowTools(workflowID, projectID, tmpTaskID, finishedFlag, userID)

                    End If

                    '    获取当前任务的转移任务和转移条件；
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                    dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    '    获取转移条件为真的转移任务集（如果转移条件为空，返回真）；
                    dsConditionTrue = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")
                    For j = 0 To dsTempTaskTransfer.Tables(0).Rows.Count - 1
                        nextTaskID = dsTempTaskTransfer.Tables(0).Rows(j).Item("next_task")
                        tmpTransCondition = Trim(dsTempTaskTransfer.Tables(0).Rows(j).Item("transfer_condition"))
                        tmpTransStatus = Trim(IIf(IsDBNull(dsTempTaskTransfer.Tables(0).Rows(j).Item("project_status")), "", dsTempTaskTransfer.Tables(0).Rows(j).Item("project_status")))

                        '判断条件是否为真
                        If CompareExpression(workflowID, projectID, tmpTaskID, ".T.", tmpTransCondition) Then

                            '构造转移条件为真的转移任务集
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

                    '    将当前任务（任务类型为AUTO）从转移活动集中删除；
                    ShiftTaskSet.Tables(0).Rows(i).Delete()
                    ShiftTaskSet.AcceptChanges()

                    '    将转移条件为真的转移任务集添加到ShiftTaskSet；
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

                    '如果找到虚任务,返回本次操作
                    Exit For

                End If


            Next
        Loop

        Return ShiftTaskSet

    End Function


    'FiliterRecedeDummyTask（ProjectID、ShiftTaskSet）
    '过滤制约活动集的虚活动，获取业务活动
    Private Function FiliterRecedeDummyTask(ByVal workflowID As String, ByVal projectID As String, ByVal ShiftTaskSet As DataSet, ByVal userID As String) As DataSet
        '①	Vtask=True
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
                '② 对于转移条件为真的制约活动集ShiftTaskSet中的每个任务TaskID
                '    在任务表获取当前任务（ProjectID、TaskID）；
                tmpTaskID = ShiftTaskSet.Tables(0).Rows(i).Item("task_id")
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & tmpTaskID & "'" & "}"

                '获取任务的活动类型；
                dsTempTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

                '异常处理  
                If dsTempTask.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempTask.Tables(0))
                    Throw wfErr
                End If

                tmpTaskType = IIf(IsDBNull(dsTempTask.Tables(0).Rows(0).Item("task_type")), "", dsTempTask.Tables(0).Rows(0).Item("task_type"))

                '如果任务类型为“AUTO”
                If tmpTaskType = "AUTO" Then

                    'Vtask= True
                    vTask = True

                    '获取当前任务的制约任务和制约条件；
                    strSql = "{project_code=" & "'" & projectID & "'" & " and next_task=" & "'" & tmpTaskID & "'" & "}"
                    dsTempTaskTransfer = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                    '将当前任务（任务类型为AUTO）从制约活动集中删除；
                    ShiftTaskSet.Tables(0).Rows(i).Delete()
                    ShiftTaskSet.AcceptChanges()

                    '将制约任务集添加到ShiftTaskSet；
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


    '刷新会议，把超过当前系统时间还没有安排项目的评审会删除
    Public Function RefreshConference()
        ''获取所有超过系统当前时间的会议
        'Dim sysDate As String = FormatDateTime(Now, DateFormat.ShortDate)
        'Dim strSql As String
        'strSql = "{conference_date<" & "'" & sysDate & "'" & "}"
        'Dim Conference As New Conference(conn, ts)
        'Dim dsConference As DataSet = Conference.GetConferenceInfo(strSql, "null")

        ''删除没有安排项目的评审会
        'Dim i As Integer
        'Dim tmpConferenceCode As String
        'Dim dsConfTrial As DataSet
        'Dim ConfTrial As New ConfTrial(conn, ts)
        'For i = 0 To dsConference.Tables(0).Rows.Count - 1
        '    tmpConferenceCode = dsConference.Tables(0).Rows(i).Item("conference_code")
        '    strSql = "{conference_code=" & "'" & tmpConferenceCode & "'" & "}"
        '    dsConfTrial = ConfTrial.GetConfTrialInfo(strSql, "null")

        '    '如果没有安排项目
        '    If dsConfTrial.Tables(0).Rows.Count = 0 Then
        '        dsConference.Tables(0).Rows(i).Delete()
        '    End If

        'Next
        'Conference.UpdateConferenceCommitteeman(dsConference)

    End Function

    '提交安排评审会议任务FinishedReviewConferencePlan（ConferenceCode）
    '为了支持在一次评审会上安排多个项目，安排评审会任务提交方法作如下调整： 
    Public Function FinishedReviewConferencePlan(ByVal ConferenceCode As String)
        '①	在Conference-Trail中获取参数ConferenceCode指定的所有项目编码；
        Dim strSql As String
        Dim i, j As Integer
        Dim newRow As DataRow
        Dim tmpProjectCode, tmpWorkflowID, tmpUserID, tmpManagerA, tmpManagerB, tmpTaskStatus As String
        Dim isHasA, isHasB As Boolean
        Dim dsTempProject, dsTempCommitteeman, dsTempAttend, dsTempMsg, dsConference As DataSet
        Dim ConfTrial As New ConfTrial(conn, ts)
        Dim Conference As New Conference(conn, ts)
        Dim CommonQuery As New CommonQuery(conn, ts)
        Dim record_person As String '会场的记录员

        Dim conferenceTime As DateTime = Now

        strSql = "{conference_code=" & "'" & ConferenceCode & "'" & "}"
        dsTempProject = ConfTrial.GetConfTrialInfo(strSql, "null")
        '②	将Conference-Committeeman表中获取的所有评委加入会议参与人集；
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
        '③	对每个项目编码
        For i = 0 To dsTempProject.Tables(0).Rows.Count - 1
            '在Project-Task-Attendee获取与项目编码一致，ReviewMeetingPlan任务状态“P”的Workflow-id和参与人；
            tmpProjectCode = dsTempProject.Tables(0).Rows(i).Item("project_code")
            isExp = IIf(IsDBNull(dsTempProject.Tables(0).Rows(i).Item("is_exp")), False, dsTempProject.Tables(0).Rows(i).Item("is_exp"))

            '如果是展期项目，提交展期得安排会议任务
            If isExp Then
                strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlanExp'}"
            Else
                strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan'}"
            End If

            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpTaskStatus = dsTempAttend.Tables(0).Rows(0).Item("task_status")

            '判断项目是否已提交过,只有未提交过的项目才可安排
            If tmpTaskStatus = "P" Then


                tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")
                tmpUserID = Trim(dsTempAttend.Tables(0).Rows(0).Item("attend_person"))

                '获取与Workflow-id、项目编码匹配的项目经理A角和B角；
                '如果项目经理A角不在会议参与人集中，将项目经理A角加入会议参与人集；
                strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
                dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

                '异常处理  
                If dsTempAttend.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr
                    wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                    Throw wfErr
                End If

                tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))

                '如果项目经理B角不在会议参与人集中，将项目经理B角加入会议参与人集；
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

                '更改处理 记录评审会结论 任务的人员为 会场的记录人员  2005-6-30 LQF add
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

                '判断项目是否已提交

                '如果是展期项目，提交展期得安排会议任务
                If isExp Then
                    strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlanExp' and task_status='P'}"
                Else
                    strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan' and task_status='P'}"
                End If

                dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                If dsTempAttend.Tables(0).Rows.Count <> 0 Then
                    If isExp Then
                        '调用FinishedTask（Workflow-id、项目编码、ReviewMeetingPlan、“ ”、参与人）。
                        finishedTask(tmpWorkflowID, tmpProjectCode, "ReviewMeetingPlanExp", "", tmpUserID)
                        setReviewConclusionCueTime(tmpProjectCode, "ReviewMeetingPlanExp", conferenceTime)
                    Else
                        '调用FinishedTask（Workflow-id、项目编码、ReviewMeetingPlan、“ ”、参与人）。
                        finishedTask(tmpWorkflowID, tmpProjectCode, "ReviewMeetingPlan", "", tmpUserID)
                        setReviewConclusionCueTime(tmpProjectCode, "ReviewMeetingPlan", conferenceTime)
                    End If


                End If
            End If

        Next

        ''④	对于会议参与人集的所有成员
        ''向消息库添加“请查看评审会议程”信息；
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")


        ''获取评审会的日期
        'strSql = "{conference_code=" & "'" & ConferenceCode & "'" & "}"
        'dsConference = Conference.GetConferenceInfo(strSql, "null")

        ''异常处理  
        'If dsConference.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsConference.Tables(0))
        '    Throw wfErr
        'End If

        'Dim tmpConfTime As String = CStr(dsConference.Tables(0).Rows(0).Item("conference_date")) & " " & dsConference.Tables(0).Rows(0).Item("start_time")

        'For i = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = "请查看" & tmpConfTime & "的评审会议议程"
        '        .Item("accepter") = dsTempCommitteeman.Tables(1).Rows(i).Item("committeeman")
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)

    End Function

    '撤消评审会议安排任务CancelReviewConferencePlan（ConferenceCode）
    Public Function CancelReviewConferencePlan(ByVal ConferenceCode As String)
        '①	在Conference-Trail中获取参数ConferenceCode指定的所有项目编码；
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

        '②	将Conference-Committeeman表获取的所有评委加入会议参与人集；
        dsTempCommitteeman = Conference.GetConferenceInfo("null", strSql)

        '③	在Conference表中获取与ConferenceCode匹配的Conference-date；
        dsTempConference = Conference.GetConferenceInfo(strSql, "null")

        '异常处理  
        If dsTempConference.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempConference.Tables(0))
            Throw wfErr
        End If

        tmpConferenceDate = dsTempConference.Tables(0).Rows(0).Item("conference_date")
        '④	对每个项目编码
        For i = 0 To dsTempProject.Tables(0).Rows.Count - 1
            tmpProjectCode = dsTempProject.Tables(0).Rows(i).Item("project_code")

            '获取该项目的workflow_id
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")

            '判断评审会是否已开过
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion' and task_status='F'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            If dsTempAttend.Tables(0).Rows.Count > 0 Then
                '抛出"不能撤销已开过的评审会"
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowRecordReviewConclusionErr()
                Throw wfErr
                Exit Function
            End If

            '把所有项目的记录评审会结论任务置空
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For j = 0 To dsTempAttend.Tables(0).Rows.Count - 1
                dsTempAttend.Tables(0).Rows(j).Item("task_status") = ""
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            ''在Project-Task-Attendee获取与项目编码一致，ReviewMeetingPlan任务状态为“F”的任务；
            'strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan' and task_status='F'" & "}"
            'dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            ''异常处理  
            'If dsTempAttend.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
            '    Throw wfErr
            'End If

            'tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")
            'tmpTaskID = Trim(dsTempAttend.Tables(0).Rows(0).Item("task_id"))

            '获取与Workflow-id、项目编码匹配的项目经理A角和B角；
            '如果项目经理A角不在会议参与人集中，将项目经理A角加入会议参与人集；
            strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
            dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))

            '如果项目经理B角不在会议参与人集中，将项目经理B角加入会议参与人集；
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

            '将任务状态置为“P”；
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='ReviewMeetingPlan'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For j = 0 To dsTempAttend.Tables(0).Rows.Count - 1
                dsTempAttend.Tables(0).Rows(j).Item("task_status") = "P"
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            ''如果记录评审会结论状态为'P' ,将任务状态置为“”；
            'strSql = "{workflow_id=" & "'" & tmpWorkflowID & "'" & " and project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordReviewConclusion'" & "}"
            'dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            'If IIf(IsDBNull(dsTempAttend.Tables(0).Rows(0).Item("task_status")), "", dsTempAttend.Tables(0).Rows(0).Item("task_status")) = "P" Then
            '    dsTempAttend.Tables(0).Rows(0).Item("task_status") = ""
            'End If
            'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        Next

        ''⑤	对于会议参与人集的所有成员
        ''向消息库添加“Conference-date评审会撤消”信息；
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")
        'For i = 0 To dsTempCommitteeman.Tables(1).Rows.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = CStr(tmpConferenceDate) & "评审会撤消"
        '        .Item("accepter") = dsTempCommitteeman.Tables(1).Rows(i).Item("committeeman")
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)

        '删除该会议的所有项目
        For i = 0 To dsTempProject.Tables(0).Rows.Count - 1
            dsTempProject.Tables(0).Rows(i).Item("conference_code") = DBNull.Value
        Next
        ConfTrial.UpdateConfTrial(dsTempProject)

        '会议删除
        dsTempConference.Tables(0).Rows(0).Delete()
        Conference.UpdateConferenceCommitteeman(dsTempConference)
    End Function

    '撤销评审会项目
    Public Function CancelReviewConferencePlanProject(ByVal projectID As String)
        Dim strSql As String
        Dim dsTempAttend As DataSet
        Dim i As Integer

        '在参与人表中将该项目的记录评审会结论任务置为""
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordReviewConclusion'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '在参与人表中将该项目的安排评审会任务置为"P"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ReviewMeetingPlan'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function

    '自动确认消息AACKMassage（ProjectID、TaskID、EmployeeID）
    Public Function AACKMassage(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String, ByVal employeeID As String)

        '①	获取ProjectID的企业名称；
        '获取项目的企业名称
        Dim strSql As String
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsProject As DataSet = project.GetProjectInfo(strSql)

        '异常处理  
        If dsProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProject.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")

        '异常处理  
        If dsCorporation.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsCorporation.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        '②	获取ProjectID、TaskID的任务名称；
        '②	获取任务ID的任务名称；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        Dim dsTempTask As DataSet = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '未找到任务则返回  
        If dsTempTask.Tables(0).Rows.Count = 0 Then
            Exit Function
        Else

            Dim tmpTaskName As String = dsTempTask.Tables(0).Rows(0).Item("task_name")

            '③	将消息库中消息内容包含企业名称和任务名称的消息确认标志改为“Y”。
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

    '（委托）获取任务参与人getTaskActor（RoleID）
    Public Function getTaskActor(ByVal roleID As String) As String
        '①	在STAFF-ROLE表获取角色ID=RoleID的staff_name和consigner;
        Dim strSql As String
        Dim i As Integer
        Dim role As New Role(conn, ts)
        strSql = "{role_id=" & "'" & roleID & "'" & "}"
        Dim dsTempStaff As DataSet = role.GetStaffRole(strSql)

        '异常处理  
        If dsTempStaff.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
            Throw wfErr
        End If

        '②	Staff_name=staff_name；
        Dim tmpStaffName As String = Trim(dsTempStaff.Tables(0).Rows(0).Item("staff_name"))
        '③	Consigner=consigner；
        Dim tmpConsigner As String = Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(0).Item("consigner")), "", dsTempStaff.Tables(0).Rows(0).Item("consigner")))
        '④	WHILE （Consigner非空）
        While tmpConsigner <> ""
            '在STAFF-ROLE表获取员工姓名为Consigner的staff_name和委托人集consignerSet；
            strSql = "{staff_name=" & "'" & tmpConsigner & "'" & "}"
            dsTempStaff = role.GetStaffRole(strSql)

            '异常处理  
            If dsTempStaff.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
                Throw wfErr
            End If

            'Staff_name=Consigner；
            tmpStaffName = tmpConsigner
            'Consigner=""
            tmpConsigner = ""
            '   For I = 0 To Number(consignerSet) - 1
            For i = 0 To dsTempStaff.Tables(0).Rows.Count - 1
                'IF consignerSet（I）IS NOT Null THEN Consigner=consigner；
                If Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(0).Item("consigner")), "", dsTempStaff.Tables(0).Rows(0).Item("consigner"))) <> "" Then
                    tmpConsigner = dsTempStaff.Tables(0).Rows(0).Item("consigner")
                End If
            Next
        End While

        '⑤	返回Staff_name；
        Return tmpStaffName
    End Function

    Public Function getTaskActor(ByVal projectID As String, ByVal taskID As String, ByVal roleID As String, ByVal branch As String) As String
        '①	在STAFF-ROLE表获取角色ID=RoleID的staff_name和consigner;
        Dim strSql As String
        Dim i As Integer
        Dim role As New Role(conn, ts)
        strSql = "{role_id=" & "'" & roleID & "'" & "}"
        Dim dsTempStaff As DataSet = role.GetStaffRole(strSql)

        '异常处理  
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
            '②	Staff_name=staff_name；
            tmpStaffName = Trim(dsTempStaff.Tables(0).Rows(i).Item("staff_name"))
            strSql = "{staff_name=" & "'" & tmpStaffName & "'" & "}"
            dsTemp = staff.FetchStaff(strSql)

            '异常处理  
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

        '如果在分支机构找到参与人
        If isFound Then

            '③	Consigner=consigner；
            Dim tmpConsigner As String = Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(iStaff).Item("consigner")), "", dsTempStaff.Tables(0).Rows(iStaff).Item("consigner")))
            '④	WHILE （Consigner非空）
            'While tmpConsigner <> ""
            '    '在STAFF-ROLE表获取员工姓名为Consigner的staff_name和委托人集consignerSet；
            '    strSql = "{staff_name=" & "'" & tmpConsigner & "'" & "}"
            '    dsTempStaff = role.GetStaffRole(strSql)

            '    '异常处理  
            '    If dsTempStaff.Tables(0).Rows.Count = 0 Then
            '        Dim wfErr As New WorkFlowErr()
            '        wfErr.ThrowNoRecordkErr(dsTempStaff.Tables(0))
            '        Throw wfErr
            '    End If

            '    'Staff_name=Consigner；
            '    tmpStaffName = tmpConsigner
            '    'Consigner=""
            '    tmpConsigner = ""
            '    '   For I = 0 To Number(consignerSet) - 1
            '    For i = 0 To dsTempStaff.Tables(0).Rows.Count - 1
            '        'IF consignerSet（I）IS NOT Null THEN Consigner=consigner；
            '        If Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(i).Item("consigner")), "", dsTempStaff.Tables(0).Rows(i).Item("consigner"))) <> "" Then
            '            tmpConsigner = dsTempStaff.Tables(0).Rows(0).Item("consigner")
            '        End If
            '    Next
            'End While

            If tmpConsigner <> "" Then
                tmpStaffName = tmpConsigner

                '将原委托人填入consinger中
                Dim tmpSrcPerson As String = Trim(IIf(IsDBNull(dsTempStaff.Tables(0).Rows(iStaff).Item("staff_name")), "", dsTempStaff.Tables(0).Rows(iStaff).Item("staff_name")))
                Dim dsConsinger As DataSet
                strSql = "{project_code='" & projectID & "' and task_id='" & taskID & "'}"
                dsConsinger = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                    dsConsinger.Tables(0).Rows(i).Item("consigner") = tmpSrcPerson
                Next

                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsConsinger)

            End If


            '⑤	返回Staff_name；
            Return tmpStaffName
        Else

            Return ""

        End If

    End Function


    'consignTask（员工、角色、委托人），由客户端使用。
    Public Function consignTask(ByVal staffID As String, ByVal roleID As String, ByVal consigner As String, ByVal isCurrent As Boolean)
        '①	在STAFF-ROLE表查找参数指定的员工、角色；
        Dim strSql As String
        strSql = "{staff_name=" & "'" & staffID & "'" & " and role_id=" & "'" & roleID & "'" & "}"
        Dim role As New Role(conn, ts)
        Dim dsTempRole As DataSet = role.GetStaffRole(strSql)


        '异常处理  
        If dsTempRole.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
            Throw wfErr
        End If

        ''如果该员工已委托了任务，提示该人已委托错误
        'Dim tmpConsigner As String = IIf(IsDBNull(dsTempRole.Tables(0).Rows(0).Item("consigner")), "", dsTempRole.Tables(0).Rows(0).Item("consigner"))

        'If tmpConsigner <> "" Then
        '    '提示该人已委托了任务错误
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowIsConsign()
        '    Throw wfErr
        '    Exit Function
        'End If

        '②	如果指定的员工角色存在
        If dsTempRole.Tables(0).Rows.Count <> 0 Then
            '在ROLE表获取员工委托角色的委托标志；
            strSql = "{role_id=" & "'" & roleID & "'" & "}"
            Dim dsTempRoleConsign As DataSet = role.FetchRole(strSql)

            '异常处理  
            If dsTempRoleConsign.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempRoleConsign.Tables(0))
                Throw wfErr
            End If


            '如果Isconsign=1 
            If IIf(IsDBNull(dsTempRoleConsign.Tables(0).Rows(0).Item("isConsign")), False, dsTempRoleConsign.Tables(0).Rows(0).Item("isConsign")) = True Then

                ' 将STAFF-ROLE表中的员工委托人consigner置为委托人；
                dsTempRole.Tables(0).Rows(0).Item("consigner") = consigner
                role.UpdateStaffRole(dsTempRole)

                '如果要委托当前任务,调用consignCurrentTask
                'If isCurrent Then
                consignCurrentTask(staffID, roleID, consigner)
                'End If

                '否则
            Else

                '提示“没有委托权限！”
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoConsignRight()
                Throw wfErr

            End If
        End If
    End Function


    'CancelconsignTask（员工、角色），由客户端使用。
    Public Function CancelconsignTask(ByVal srcPerson As String, ByVal staffID As String, ByVal roleID As String, ByVal isCurrent As Boolean)
        '①	在STAFF-ROLE表查找与参数员工、角色匹配的记录；
        Dim strSql As String
        strSql = "{staff_name=" & "'" & srcPerson & "'" & " and role_id=" & "'" & roleID & "'" & "}"
        Dim role As New Role(conn, ts)
        Dim dsTempRole As DataSet = role.GetStaffRole(strSql)

        '异常处理  
        If dsTempRole.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
            Throw wfErr
        End If


        '②	如果查找记录存在，将consigner置为NULL；
        If dsTempRole.Tables(0).Rows.Count <> 0 Then
            If dsTempRole.Tables(0).Rows(0).Item("consigner") Is DBNull.Value Then
                '提示“无受托人!”
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoConsigner()
                Throw wfErr
            Else
                dsTempRole.Tables(0).Rows(0).Item("consigner") = DBNull.Value

                '如果要撤销当前任务的委托,调用CancelconsignCurrentTask
                'If isCurrent Then
                CancelconsignCurrentTask(staffID, roleID, srcPerson)
                ' End If

            End If
        End If

        role.UpdateStaffRole(dsTempRole)

    End Function

    '委托当前任务
    Private Function consignCurrentTask(ByVal staffID As String, ByVal roleID As String, ByVal consigner As String)
        ''将参与人表中任务为P，ROLEID=委托角色的任务参与人改为最终参与人
        'Dim strSql As String
        'Dim i As Integer
        'strSql = "{role_id='" & roleID & "' and attend_person='" & staffID & "' and task_status='P'}"
        'Dim dsTempAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        ''如果当前任务为资产评估CapitialEvaluated，需将反担保物的评估师置为受托人
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

        ''记录原委托人
        '' 获取原参与人
        '' 记录原参与人
        'Dim dsSrcPerson As String
        'For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
        '    dsSrcPerson = dsTempAttend.Tables(0).Rows(i).Item("attend_person")
        '    dsTempAttend.Tables(0).Rows(i).Item("consigner") = dsSrcPerson
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        ''将参与人改为受托人
        ''Dim tmpLastAttend As String = getTaskActor(roleID)
        'For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
        '    dsTempAttend.Tables(0).Rows(i).Item("attend_person") = consigner
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '将参与人表中任务为P，ROLEID=委托角色的任务参与人改为最终参与人
        Dim strSql As String
        Dim i As Integer
        strSql = "{role_id='" & roleID & "' and attend_person='" & staffID & "'}"
        Dim dsTempAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)


        '如果当前任务为资产评估CapitialEvaluated，需将反担保物的评估师置为受托人
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

        '记录原委托人
        ' 获取原参与人
        ' 记录原参与人
        '将参与人改为受托人
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("attend_person") = consigner
            dsTempAttend.Tables(0).Rows(i).Item("consigner") = staffID
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function

    '撤销当前任务委托
    Private Function CancelconsignCurrentTask(ByVal staffID As String, ByVal roleID As String, ByVal srcPerson As String)
        ''将参与人表中任务为P，ROLEID=委托角色的任务参与人改为原角色员工
        'Dim strSql As String
        'Dim i As Integer
        'Dim dsTempAttend As DataSet

        ''如果当前任务为资产评估CapitialEvaluated，需将反担保物的评估师置为原委托人
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
        ''将参与人改为原委托人
        'For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1

        '    '获取改任务的原委托人
        '    dsSrcPerson = IIf(IsDBNull(dsTempAttend.Tables(0).Rows(i).Item("consigner")), "", dsTempAttend.Tables(0).Rows(i).Item("consigner"))

        '    '如果原委托人非空
        '    If dsSrcPerson <> "" Then
        '        dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsSrcPerson
        '        dsTempAttend.Tables(0).Rows(i).Item("consigner") = ""
        '    End If
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '将参与人表中任务为P，ROLEID=委托角色的任务参与人改为原角色员工
        Dim strSql As String
        Dim i As Integer
        Dim dsTempAttend As DataSet

        '如果当前任务为资产评估CapitialEvaluated，需将反担保物的评估师置为原委托人
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
        '将参与人改为原委托人
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1

            '获取改任务的原委托人
            dsSrcPerson = IIf(IsDBNull(dsTempAttend.Tables(0).Rows(i).Item("consigner")), "", dsTempAttend.Tables(0).Rows(i).Item("consigner"))

            '如果原委托人非空
            If dsSrcPerson <> "" Then
                dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsSrcPerson
                dsTempAttend.Tables(0).Rows(i).Item("consigner") = ""
            End If
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function


    '格式：AddTaskTrackRecord（ProjectID,Workflow_id,TaskID,StartupTask）
    Public Function AddTaskTrackRecord(ByVal workflowID As String, ByVal projectID As String, ByVal finishedTaskID As String, ByVal StartupTaskID As String)
        Dim strSql As String
        Dim i As Integer

        '①	创建Project_Track对象;
        Dim dsProjectTrack As DataSet
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo("null")

        '②	将参数值分别赋予新增对象的Project_Code、Workflow_id、FinishedTask、StartupTask属性,Status属性置为“P”；；
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

        '③	将所有满足Project_Code= ProjectID 、StartupTask= TaskID、Status=“P”的Project_Track对象的Status属性置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and StartupTask=" & "'" & finishedTaskID & "'" & " and isnull(status,'')='P'}"
        dsProjectTrack = WfProjectTrack.GetWfProjectTrackInfo(strSql)
        For i = 0 To dsProjectTrack.Tables(0).Rows.Count - 1
            dsProjectTrack.Tables(0).Rows(i).Item("Status") = "F"
        Next
        WfProjectTrack.UpdateWfProjectTrack(dsProjectTrack)

    End Function


    '提交签约计划
    Public Function FinishedSignaturePlan(ByVal SignaturePlanCode As Integer)
        Dim i, j As Integer
        Dim strSql As String
        Dim ProjectSignature As New ProjectSignature(conn, ts)
        Dim SignaturePlan As New SignaturePlan(conn, ts)
        Dim dsProjectSignature, dsTempAttend, dsTempMsg, dsSignaturePlan, dsTimingTask As DataSet
        Dim tmpProjectCode, tmpWorkflowID, tmpUserID, tmpTaskStatus, tmpManagerA, tmpManagerB, tmpMinister, tmpManagerLaw, tmpDirector As String
        Dim CommonQuery As New CommonQuery(conn, ts)

        '获取该签约计划的所有项目
        strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        dsProjectSignature = ProjectSignature.GetProjectSignatureInfo(strSql)

        strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        dsSignaturePlan = SignaturePlan.GetSignaturePlanInfo(strSql)

        '异常处理  
        If dsSignaturePlan.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsSignaturePlan.Tables(0))
            Throw wfErr
        End If

        Dim tmpSignaturePlanDate As DateTime = dsSignaturePlan.Tables(0).Rows(0).Item("signature_plan_date")


        For i = 0 To dsProjectSignature.Tables(0).Rows.Count - 1
            tmpProjectCode = dsProjectSignature.Tables(0).Rows(i).Item("project_code")

            '在Project-Task-Attendee获取与项目编码一致，PlanSignature任务状态为“P”的Workflow-id和参与人；
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='PlanSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpTaskStatus = dsTempAttend.Tables(0).Rows(0).Item("task_status")

            '判断项目是否已提交过,只有未提交过的项目才可安排
            If tmpTaskStatus = "P" Then


                tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")
                tmpUserID = Trim(dsTempAttend.Tables(0).Rows(0).Item("attend_person"))

                '调用FinishedTask（Workflow-id、项目编码、PlanSignature、“ ”、参与人）。
                finishedTask(tmpWorkflowID, tmpProjectCode, "PlanSignature", "", tmpUserID)

                '将登记签约定时任务的状态置为"P",开始时间置为签约计划的时间
                strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'}"
                dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                For j = 0 To dsTimingTask.Tables(0).Rows.Count - 1
                    dsTimingTask.Tables(0).Rows(j).Item("status") = "P"
                    dsTimingTask.Tables(0).Rows(j).Item("start_time") = tmpSignaturePlanDate
                Next
                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

            End If
        Next

        '获取与Workflow-id、项目编码匹配的项目经理A角和B角,中心主任，风险部长，法物经理；
        '如果项目经理A角不在会议参与人集中，将项目经理A角加入会议参与人集；
        strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
        dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

        '异常处理  
        If dsTempAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
            Throw wfErr
        End If

        Dim ArrAttend As New ArrayList
        tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))
        ArrAttend.Add(tmpManagerA)

        '如果项目经理B角不在会议参与人集中，将项目经理B角加入会议参与人集；
        tmpManagerB = Trim(dsTempAttend.Tables(0).Rows(0).Item("25"))
        ArrAttend.Add(tmpManagerB)
        tmpMinister = Trim(dsTempAttend.Tables(0).Rows(0).Item("31"))
        ArrAttend.Add(tmpMinister)
        tmpManagerLaw = Trim(dsTempAttend.Tables(0).Rows(0).Item("33"))

        '风险部长和法物经理不为同一人
        If tmpManagerLaw <> tmpMinister Then
            ArrAttend.Add(tmpManagerLaw)
        End If

        tmpDirector = getTaskActor("01")
        ArrAttend.Add(tmpDirector)

        ''⑤	对于签约参与人集的所有成员
        ''向消息库添加“Signature-plan-date”出席签约信息；
        'Dim newRow As DataRow
        'dsTempMsg = WfProjectMessages.GetWfProjectMessagesInfo("null")
        'For i = 0 To ArrAttend.Count - 1
        '    newRow = dsTempMsg.Tables(0).NewRow
        '    With newRow
        '        .Item("message_content") = CStr(tmpSignaturePlanDate) & "请出席签约"
        '        .Item("accepter") = ArrAttend(i)
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)


    End Function

    '撤销签约
    Public Function CancelSignaturePlan(ByVal SignaturePlanCode As Integer)
        Dim i, j As Integer
        Dim strSql As String
        Dim ProjectSignature As New ProjectSignature(conn, ts)
        Dim SignaturePlan As New SignaturePlan(conn, ts)
        Dim dsProjectSignature, dsTempAttend, dsTempMsg, dsSignaturePlan, dsTimingTask As DataSet
        Dim tmpProjectCode, tmpWorkflowID, tmpUserID, tmpTaskID, tmpManagerA, tmpManagerB, tmpMinister, tmpManagerLaw, tmpDirector As String
        Dim CommonQuery As New CommonQuery(conn, ts)

        '获取该签约计划的所有项目
        strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        dsProjectSignature = ProjectSignature.GetProjectSignatureInfo(strSql)

        '④	对每个项目编码
        For i = 0 To dsProjectSignature.Tables(0).Rows.Count - 1
            tmpProjectCode = dsProjectSignature.Tables(0).Rows(i).Item("project_code")

            '获取该项目的workflow_id
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")

            '判断项目是否已签约
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature' and task_status='F'}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            If dsTempAttend.Tables(0).Rows.Count > 0 Then
                '抛出"不能撤销已签约的计划"
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowRecordSignatureErr()
                Throw wfErr
                Exit Function
            End If

            '把所有项目的登记签约任务置空
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For j = 0 To dsTempAttend.Tables(0).Rows.Count - 1
                dsTempAttend.Tables(0).Rows(i).Item("task_status") = ""
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            '在Project-Task-Attendee获取与项目编码一致，PlanSignature任务状态为“F”的任务；
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='PlanSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            tmpWorkflowID = dsTempAttend.Tables(0).Rows(0).Item("workflow_id")

            '将任务状态置为“P”；
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='PlanSignature'" & "}"
            dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTempAttend.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempAttend.Tables(0))
                Throw wfErr
            End If

            dsTempAttend.Tables(0).Rows(0).Item("task_status") = "P"
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

            '将登记签约定时任务的状态置为DBNULL,开始时间置为DBNULL
            strSql = "{project_code=" & "'" & tmpProjectCode & "'" & " and task_id='RecordSignature'}"
            dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
            For j = 0 To dsTimingTask.Tables(0).Rows.Count - 1
                dsTimingTask.Tables(0).Rows(j).Item("status") = DBNull.Value
                'dsTimingTask.Tables(0).Rows(j).Item("start_time") = DBNull.Value
            Next
            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

        Next

        '获取与Workflow-id、项目编码匹配的项目经理A角和B角,中心主任，风险部长，法物经理；
        '如果项目经理A角不在会议参与人集中，将项目经理A角加入会议参与人集；
        strSql = "{ProjectCode=" & "'" & tmpProjectCode & "'" & "}"
        dsTempAttend = CommonQuery.GetProjectInfoEx(strSql)

        Dim ArrAttend As New ArrayList
        tmpManagerA = Trim(dsTempAttend.Tables(0).Rows(0).Item("24"))
        ArrAttend.Add(tmpManagerA)

        '如果项目经理B角不在会议参与人集中，将项目经理B角加入会议参与人集；
        tmpManagerB = Trim(dsTempAttend.Tables(0).Rows(0).Item("25"))
        ArrAttend.Add(tmpManagerB)
        tmpMinister = Trim(dsTempAttend.Tables(0).Rows(0).Item("31"))
        ArrAttend.Add(tmpMinister)
        tmpManagerLaw = Trim(dsTempAttend.Tables(0).Rows(0).Item("33"))

        '风险部长和法物经理不为同一人
        If tmpManagerLaw <> tmpMinister Then
            ArrAttend.Add(tmpManagerLaw)
        End If

        tmpDirector = getTaskActor("01")
        ArrAttend.Add(tmpDirector)

        ''⑤	对于签约参与人集的所有成员
        ''向消息库添加“Signature-plan-date签约撤消”信息；
        'Dim newRow As DataRow
        'strSql = "{signature_plan_code=" & SignaturePlanCode & "}"
        'dsSignaturePlan = SignaturePlan.GetSignaturePlanInfo(strSql)

        ''异常处理  
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
        '        .Item("message_content") = CStr(tmpSignaturePlanDate) & "签约撤消"
        '        .Item("accepter") = ArrAttend(i)
        '        .Item("send_time") = Now
        '        .Item("is_affirmed") = "N"
        '    End With
        '    dsTempMsg.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectMessages.UpdateWfProjectMessages(dsTempMsg)


        '删除该会议的所有项目
        For i = 0 To dsProjectSignature.Tables(0).Rows.Count - 1
            dsProjectSignature.Tables(0).Rows(i).Item("signature_plan_code") = DBNull.Value
        Next
        ProjectSignature.UpdateProjectSignature(dsProjectSignature)

        '会议删除
        dsSignaturePlan.Tables(0).Rows(0).Delete()
        SignaturePlan.UpdateSignaturePlan(dsSignaturePlan)
    End Function


    '撤销签约计划的项目
    Public Function CancelSignaturePlanProject(ByVal projectID As String)

        Dim strSql As String
        Dim dsTempAttend, dsTimingTask As DataSet
        Dim i As Integer

        '在参与人表中将该项目的记录签约计划任务置为""
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordSignature'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '在参与人表中将该项目的安排签约任务置为"P"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='PlanSignature'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '将登记签约定时任务的状态置为DBNULL
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordSignature'}"
        dsTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTimingTask.Tables(0).Rows.Count - 1
            dsTimingTask.Tables(0).Rows(i).Item("status") = DBNull.Value
            'dsTimingTask.Tables(0).Rows(i).Item("start_time") = DBNull.Value
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

    End Function

    '重新上会
    Public Function ReMeetingPlan(ByVal projectID As String)
        Dim strSql As String
        Dim i As Integer
        Dim dsTemp As DataSet

        '2010-11-03 yjf add 如果法务在记录评审会结论，项目经理不可以重新上会
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordReviewConclusion'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("task_status")), "", dsTemp.Tables(0).Rows(0).Item("task_status")) = "P" Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowMustRecord()
            Throw wfErr
        End If

        '将当前处理任务关闭
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_status='P'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            '
            'qxd modify 2004-7-22 
            '
            '重新上会时将项目当前任务状态为P的任务置为A(保留任务状态)。
            '在记录评审会结论界面增加“恢复项目状态”选项，选择恢复项目状态，
            '将项目任务状态为A的任务状态置回P；否则，项目按流程流转。
            '建议将是否是最终结论选项改为“恢复项目状态”。

            'dsTemp.Tables(0).Rows(i).Item("task_status") = ""
            dsTemp.Tables(0).Rows(i).Item("task_status") = "A"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '2009-4-20 yjf add
        '如果业务品种为项下保函，将项目评审置起
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

            '将ProjectProbe任务的后继任务（提交调研结论）置为"P"
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
            '将安排评审会任务状态置为空（否则重新上会其后置的任务有可能因为其状态为完成，而启动）
            strSql = "{project_code='" & projectID & "' and task_id='ReviewMeetingPlan'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("task_status") = ""
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)
        End If

        '2005-10-26 yjf add 
        '将项目的阶段改为评审，状态为重新上会
        '将项目原有状态记录在项目表中，以便恢复项目状态时恢复其阶段和状态
        Dim ojbProject As New Project(conn, ts)
        strSql = "{project_code='" & projectID & "'}"
        dsTemp = ojbProject.GetProjectInfo(strSql)
        Dim tmpPhase, tmpStatus As String
        tmpPhase = dsTemp.Tables(0).Rows(0).Item("phase")
        tmpStatus = dsTemp.Tables(0).Rows(0).Item("status")

        dsTemp.Tables(0).Rows(0).Item("phase") = "评审"
        dsTemp.Tables(0).Rows(0).Item("status") = "重新上会"


        dsTemp.Tables(0).Rows(0).Item("origPhase") = tmpPhase
        dsTemp.Tables(0).Rows(0).Item("origStatus") = tmpStatus

        ojbProject.UpdateProject(dsTemp)

    End Function

    '获得某项目的某个任务的后继任务的task_id

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

    '多次呈请放款
    Public Function ReLoanApplication(ByVal projectID As String)
        Dim strSql As String
        Dim i As Integer
        Dim dsTemp As DataSet

        '将呈请放款置为"P"
        ''strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='LoanApplication'}"
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckedSignature'}"
        'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    dsTemp.Tables(0).Rows(i).Item("task_status") = "P"
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '2004-3-18 start
        '通过调用提交“登记签约”（RecordSignature）

        Dim strUser As String
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordSignature'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If dsTemp.Tables(0).Rows.Count > 0 Then
            strUser = dsTemp.Tables(0).Rows(0).Item("attend_person")
        End If

        finishedTask("", projectID, "RecordSignature", "", strUser)
        'end 
    End Function

    '项目拆分
    Public Function SplitPrjoect(ByVal fatherProjectID As String, ByVal sonProjectID As String)
        Dim strSql As String
        Dim i As Integer
        Dim newRow As DataRow
        Dim dsTempFather, dsTempSon As DataSet
        Dim dsAttend As DataSet
        Dim strWorkflowID As String

        ''获取父项目申请品种的流程ID
        'strSql = "{project_code='" & fatherProjectID & "' and task_id='RecordReviewConclusion'}"
        'dsTempFather = WfProjectTask.GetWfProjectTaskInfo(strSql)

        ''异常处理  
        'If dsTempFather.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoRecordkErr(dsTempFather.Tables(0))
        '    Throw wfErr
        'End If

        '2011-9-1 yjf edit 获取拆分项目的流程ID

        Dim dsTempProject As DataSet = project.GetProjectInfo("{project_code='" & sonProjectID & "'}")

        '异常处理  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        strWorkflowID = dsTempProject.Tables(0).Rows(0).Item("split_workflow_id")


        Dim dsTemp, dsTemplate As DataSet

        Dim strWorkflow As String = "workflow_id=" & "'" & strWorkflowID & "'"

        '任务模板
        dsTemplate = GetWfProjectTaskTemplateInfo("task_template", strWorkflow)
        dsTemp = WfProjectTask.GetWfProjectTaskInfo("null")

        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1


            newRow = dsTemp.Tables(0).NewRow()
            With newRow

                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '将所有添加任务的工作流ID置为项目编码
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

        '角色模板
        dsTemplate = GetWfProjectTaskTemplateInfo("task_role_template", strWorkflow)
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo("null")

        '将角色模板添加到角色表中
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '将所有添加任务的工作流ID置为项目编码
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '3、转移条件模版
        dsTemplate = GetWfProjectTaskTemplateInfo("task_transfer_template", strWorkflow)
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")

        '将转移条件模版添加到转移条件表中
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '将所有添加任务的工作流ID置为项目编码
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("next_task") = dsTemplate.Tables(0).Rows(i).Item("next_task")
                .Item("transfer_condition") = dsTemplate.Tables(0).Rows(i).Item("transfer_condition")
                .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                .Item("isItem") = dsTemplate.Tables(0).Rows(i).Item("isItem")
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)


        '4、定时任务模板
        dsTemplate = GetWfProjectTaskTemplateInfo("timing_task_template", strWorkflow)
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")

        '将任务模板添加到任务模板实例表中
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = sonProjectID    '将所有添加任务的工作流ID置为项目编码
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


        '将项目经理A、B角为空的员工置为项目经理。
        Dim tmpManagerA, tmpManagerB As String
        strSql = "select nowManagerA,nowManagerB from queryProjectInfo where projectCode='" & fatherProjectID & "'"
        dsTempFather = commQuery.GetCommonQueryInfo(strSql)

        '异常处理  
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

        '2011-9-1 yjf add 设置拆分项目的评审会记录员
        Dim strFatherAttendee, strFatherWorkflowID As String
        strSql = "{project_code=" & "'" & fatherProjectID & "'" & " and task_id='" & "RecordReviewConclusion" & "'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
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
        '    strMeetingRecordPerson = "陈凤丹"
        'Else
        '    If strFatherWorkflowID = "31" Or strFatherWorkflowID = "32" Or strFatherWorkflowID = "33" Or strFatherWorkflowID = "34" Then
        '        strMeetingRecordPerson = "徐胜男"
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
            '异常处理  
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
        '复制项目参与人表：project_responsible

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

    '发送消息
    Private Function AddMsg(ByVal projectID As String, ByVal taskID As String, ByVal msg As String, ByVal accepterID As String, ByVal respsonserID As String)

        '⑥	在消息库添加消息（提示消息、员工、“N”）；
        '  组合消息内容
        Dim msgContent As String
        msgContent = respsonserID & " " & msg
        Dim dsTempTaskMessages As DataSet
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")
        Dim newRow As DataRow = dsTempTaskMessages.Tables(0).NewRow
        ' 添加消息
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

    '项目暂停时发送消息给上级主管:中心主任(01),担保部长(21),风险部长(31)
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
                strCorporation = dsCorporation.Tables(0).Rows(0).Item("corporation_name") & "项目(" & projectID & "):"
            Else
                strCorporation = "项目 " & projectID
            End If
        End If

        If Not ds Is Nothing Then
            count = ds.Tables(0).Rows.Count
            If count > 0 Then
                For i = 0 To count - 1
                    strStaff = ds.Tables(0).Rows(i).Item("staff_name")
                    AddMsg(projectID, "", "项目暂停 " & delayDay & " (天)", strStaff, strCorporation)
                Next
            End If
        End If
    End Sub

    '判断任务状态是否等于“P”
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

    '设置记录评审会结论的提示时间=会议时间+间隔时间
    Private Sub setReviewConclusionCueTime(ByVal projectID As String, ByVal taskID As String, ByVal conferenceTime As DateTime)
        Dim strSql As String
        Dim i As Integer

        '在定时活动查找与当前任务ID匹配的定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and type='A'}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)


        Dim tmpTimeLimit As Integer

        '将记录评审会结论定时任务的开始时间置为任务的会议时间＋提示间隔

        Dim newRow As DataRow
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            '提示期限
            tmpTimeLimit = dsTempTimingTask.Tables(0).Rows(i).Item("time_limit")
            newRow = dsTempTimingTask.Tables(0).Rows(i)
            With newRow
                .Item("start_time") = DateAdd(DateInterval.Hour, tmpTimeLimit * 24, conferenceTime) '启动时间＝会议时间＋提示间隔
            End With
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)
    End Sub

    '更新流程
    Public Function updateProcess()

        '1、获得isLiving=1的项目编码集合 
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

        '2、
        Dim projectCode, workFlowID As String

        For i = 0 To count - 1
            projectCode = ds.Tables(0).Rows(i).Item("project_code")
            workFlowID = ds.Tables(0).Rows(i).Item("workflow_id")
            CommonQuery.PUpdateProcess(projectCode, workFlowID)
        Next
    End Function

    Public Function updateProcess(ByVal ProjectCode As String)

        '1、获得isLiving=1的项目编码集合 
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

        '2、
        Dim workFlowID As String

        For i = 0 To count - 1
            'projectCode = ds.Tables(0).Rows(i).Item("project_code")
            workFlowID = ds.Tables(0).Rows(i).Item("workflow_id")
            CommonQuery.PUpdateProcess(ProjectCode, workFlowID)
        Next
    End Function
End Class
