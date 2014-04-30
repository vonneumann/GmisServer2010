Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'删除定时任务
Public Class ImplDeleteTimingTask
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义工作流对象引用
    Private WorkFlow As WorkFlow

    Private TimingServer As TimingServer

    Private workLog As workLog

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

        '实例化定时任务对象
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        '实例化工作流对象
        WorkFlow = New WorkFlow(conn, ts)

        workLog = New WorkLog(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)

        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTempTaskAttendee As DataSet
        Dim i As Integer

        '①将还款登记任务（TID=RefundRecord）状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundRecord'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RefundRecord", userID)

        ' '2013-11-30 yjf add 将收取委贷利息任务（TID=LoanInterestFee）状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='LoanInterestFee'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "LoanInterestFee", userID)


        '②	将登记保后调查记录任务（TID=RecordProjectProcess）状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectProcess", userID)

        '③	将审核保后调查记录任务（TID=CheckProjectProcess）状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "CheckProjectProcess", userID)

        '③	将评价项目进展任务（TID=AppraiseProjectProcess）状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='AppraiseProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "AppraiseProjectProcess", userID)

        '将登记保后活动记录任务（TID=RecordProjectTraceInfo）状态置为“F”
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id ='RecordProjectTraceInfo'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)


        '将复核保后活动记录任务（TID=CheckProjectTraceInfo）状态置为“F”
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckProjectTraceInfo'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)

        '将登记还款证明书任务（TID=RecordRefundCertificate）状态置为“F”
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordRefundCertificate'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)

        'qxd start
        '将登记项目清算信息任务（TID=RecordRefundCertificate）状态置为“F”
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundDebtInfo_claim'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)

        '将登记追偿活动任务（TID=RecordRefundCertificate）状态置为“F”
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundDebtTrailRecord_claim'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)
        'end qxd

        '④	在定时任务表中删除与（模板ID、项目编码）匹配的定时类型为“T”或“P”的所有定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and status in ('T','P')}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '将登记逾期信息,记录逾期活动,登记代偿信息,登记代偿活动任务（TID=RecordProjectTraceInfo）状态置为“”

        '2007-07-12 yjf add
        '逾期分配的分配法物经理也要关闭
        '归档项目资料关闭
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('SubmissionProjectArchives','OverdueRecord','Overdue_AssignBarrister','OverdueTrailRecord','RefundDebtInfo','RefundDebtTrailRecord')}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "OverdueRecord", userID)
        WorkFlow.AACKMassage(workFlowID, projectID, "RefundDebtInfo", userID)
        WorkFlow.AACKMassage(workFlowID, projectID, "OverdueTrailRecord", userID)
        WorkFlow.AACKMassage(workFlowID, projectID, "RefundDebtTrailRecord", userID)

        'sendMesgToManager(workFlowID, projectID)



        '2011-5-20 YJF ADD 
        '设置注销人员
        '获取项目经理所在的部门
        strSql = "select dept_name from staff where staff_name='" & userID & "'"
        Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        strSql = "select staff_name from staff where  isnull(unchain_department_list,'') like '%" & strDeptName & "%'"
        Dim dsTemp2 As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        Dim strPerson As String
        If dsTemp2.Tables(0).Rows.Count <> 0 Then
            strPerson = dsTemp2.Tables(0).Rows(0).Item("staff_name")
        End If


        '设置本项目的注销人员
        strSql = "{project_code='" & projectID & "' and role_id='56'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
            drTemp.Item("attend_person") = strPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)



        ''2008-5-13 YJF ADD 
        ''设置注销文件经办人
        ''获取项目经理所在的部门
        'strSql = "select dept_name from staff where staff_name='" & userID & "'"
        'Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        ''异常处理  
        'If dsTemp.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '    Throw wfErr
        'End If

        'Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        ''获取所有注销文件经办人
        'strSql = "select staff_name from staff_role where role_id='56'"
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

        ''异常处理  
        'If dsTemp.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoStaffRole()
        '    Throw wfErr
        'End If

        'Dim j As Integer
        'Dim strStaff, strConsigner As String
        'Dim dsTemp2, dsTemp3 As DataSet
        'Dim isFound As Boolean
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    strStaff = dsTemp.Tables(0).Rows(i).Item("staff_name")
        '    strSql = "select staff_name from staff where staff_name='" & strStaff & "' and isnull(unchain_department_list,'') like '%" & strDeptName & "%'"
        '    dsTemp2 = CommonQuery.GetCommonQueryInfo(strSql)
        '    If dsTemp2.Tables(0).Rows.Count <> 0 Then
        '        isFound = True
        '        '判断是否有设置委托，如果有则由委托人处理
        '        strSql = "select * from staff_role where role_id='56' and staff_name='" & strStaff & "'"
        '        dsTemp3 = CommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp3.Tables(0).Rows.Count <> 0 Then
        '            strConsigner = Trim(IIf(IsDBNull(dsTemp3.Tables(0).Rows(0).Item("consigner")), "", dsTemp3.Tables(0).Rows(0).Item("consigner")))
        '            If strConsigner <> "" Then
        '                strStaff = strConsigner
        '            End If
        '        End If
        '        Exit For
        '    End If
        'Next

        ''异常处理  
        'If isFound = False Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoStaffRole()
        '    Throw wfErr
        'End If


        ''设置本项目的注销文件经办人
        'strSql = "{project_code='" & projectID & "' and role_id='56'}"
        'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For j = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
        '    dsTempTaskAttendee.Tables(0).Rows(j).Item("attend_person") = strStaff
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

    End Function

    Private Function sendMesgToManager(ByVal workFlowID As String, ByVal projectID As String)
        Dim strRoleID As String = "31" '风险管理部长的role_id
        Dim count As Integer
        Dim strAttend As String
        Dim strSql As String
        Dim dsTemp As DataSet

        strSql = "{project_code='" & projectID & "' and role_id='" & strRoleID & "'}"
        dsTemp = workLog.GetWorkLogInfo(strSql)
        If Not dsTemp Is Nothing Then
            count = dsTemp.Tables(0).Rows.Count
            If count > 0 Then
                strAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
                TimingServer.AddMsg(workFlowID, projectID, "RecordRefundCertificate", strAttend, "27", "N") '27:是Message_template中的id
            End If
        End If

    End Function
End Class
