Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplAddRefundPlanExp
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
        Dim dsTempAccountDetail, dsTempTimingTask, dsTempAttend, dsTempTracePlan As DataSet
        Dim tmpReturnStartDate, tmpDeadlineDate, tmpTraceDate As DateTime
        Dim newRow As DataRow
        Dim CommonQuery As New CommonQuery(conn, ts)

        '①	在Project-account-detial表获取item-code=002、item-type=34的所有还款开始时间date。
        strSql = "SELECT DISTINCT project_code, date" & _
                 " FROM dbo.project_account_detail" & _
                 " WHERE item_code='002' and item_type='34'" & _
                 " and project_code = " & "'" & projectID & "'"
        dsTempAccountDetail = CommonQuery.GetCommonQueryInfo(strSql)

        '②	为每个还款计划在定时任务表插入还款提示任务

        '先删除原有的还款提示
        strSql = "{project_code='" & projectID & "' and task_id='RefundRecord'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '模板ID= workflow_id；
        '项目ID= project_id；
        '任务ID= task_id；
        '角色ID=24；
        '        类型 = "P"
        '开始时间=date
        '        间隔 = 0
        '将定时任务状态置为“P”；
        For i = 0 To dsTempAccountDetail.Tables(0).Rows.Count - 1
            tmpReturnStartDate = dsTempAccountDetail.Tables(0).Rows(i).Item("date")
            newRow = dsTempTimingTask.Tables(0).NewRow
            With newRow
                .Item("workflow_id") = workFlowID
                .Item("project_code") = projectID
                .Item("task_id") = "RefundRecord"
                .Item("workflow_id") = workFlowID
                .Item("role_id") = "24"
                .Item("type") = "P"
                .Item("start_time") = tmpReturnStartDate
                .Item("status") = "P"
                .Item("time_limit") = 0
                .Item("distance") = 0
                .Item("message_id") = 15
            End With
            dsTempTimingTask.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        ''添加保后跟踪记录的定时提示任务
        'dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")
        'strSql = "{project_code=" & "'" & projectID & "'" & "}"
        'dsTempTracePlan = TracePlan.GetTracePlanInfo(strSql)
        'For i = 0 To dsTempTracePlan.Tables(0).Rows.Count - 1
        '    tmpTraceDate = dsTempTracePlan.Tables(0).Rows(i).Item("trace_date")
        '    newRow = dsTempTimingTask.Tables(0).NewRow
        '    With newRow
        '        .Item("workflow_id") = workFlowID
        '        .Item("project_code") = projectID
        '        .Item("task_id") = "RecordProjectTraceInfo"
        '        .Item("workflow_id") = workFlowID
        '        .Item("role_id") = "24"
        '        .Item("type") = "P"
        '        .Item("start_time") = tmpTraceDate
        '        .Item("status") = "P"
        '        .Item("time_limit") = 0
        '        .Item("distance") = 0
        '        .Item("message_id") = 15
        '    End With
        '    dsTempTimingTask.Tables(0).Rows.Add(newRow)
        'Next
        'WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '③	在定时任务表匹配workflow_id,project_id,task_id=OverdueRecord定时任务；
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('OverdueRecord','OverdueRecordMsg')}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('OverdueRecord') and type='T'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''异常处理  
        'If dsTempTimingTask.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsTempTimingTask.Tables(0))
        '    Throw wfErr
        'End If


        '④	获取担保还款截止日期；
        'strSql = "{project_code=" & "'" & projectID & "'" & "}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and  isnull(end_date,'')<>''}"
        Dim LoanNotice As New LoanNotice(conn, ts)
        Dim dsTempLoanNotice As DataSet = LoanNotice.GetLoanNoticeInfo(strSql)

        '异常处理  
        If dsTempLoanNotice.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempLoanNotice.Tables(0))
            Throw wfErr
        End If

        tmpDeadlineDate = dsTempLoanNotice.Tables(0).Rows(0).Item("end_date")
        '逾期时间为截至日期+1天
        tmpDeadlineDate = DateAdd(DateInterval.Day, 1, tmpDeadlineDate)

        '⑤	将登记项目逾期提示任务的开始时间置为还款截止日期；将任务状态置为“P”；
        Dim j, count As Integer
        count = dsTempTimingTask.Tables(0).Rows.Count
        If count > 0 Then
            For j = 0 To count - 1
                dsTempTimingTask.Tables(0).Rows(j).Item("start_time") = tmpDeadlineDate
                dsTempTimingTask.Tables(0).Rows(j).Item("status") = "P"
            Next
        End If
        'dsTempTimingTask.Tables(0).Rows(0).Item("start_time") = tmpDeadlineDate
        'dsTempTimingTask.Tables(0).Rows(0).Item("status") = "P"

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '项目在还款到期日的第二天没有提交还款证明书任务时,向项目经理发送项目逾期信息
        '如果还款到期日三天后没有提交还款证明书任务,向主任、风险主管、担保部长（具体以用户在定时服务模版定义角色为准）发送项目逾期消息

        '先删除原有的逾期提示
        strSql = "{project_code='" & projectID & "' and task_id='OverdueRecord' and type='A'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)


        strSql = "{project_code is null}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "OverdueRecord"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "24"
            .Item("type") = "A"
            .Item("start_time") = tmpDeadlineDate
            .Item("status") = "P"
            .Item("time_limit") = 0
            .Item("distance") = 0
            .Item("message_id") = 9
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "OverdueRecord"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "01"
            .Item("type") = "A"
            .Item("start_time") = DateAdd(DateInterval.Day, 2, tmpDeadlineDate)
            .Item("status") = "P"
            .Item("time_limit") = 0
            .Item("distance") = 0
            .Item("message_id") = 9
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "OverdueRecord"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "21"
            .Item("type") = "A"
            .Item("start_time") = DateAdd(DateInterval.Day, 2, tmpDeadlineDate)
            .Item("status") = "P"
            .Item("time_limit") = 0
            .Item("distance") = 0
            .Item("message_id") = 9
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "OverdueRecord"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "31"
            .Item("type") = "A"
            .Item("start_time") = DateAdd(DateInterval.Day, 2, tmpDeadlineDate)
            .Item("status") = "P"
            .Item("time_limit") = 0
            .Item("distance") = 0
            .Item("message_id") = 9
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "OverdueRecord"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "32"
            .Item("type") = "A"
            .Item("start_time") = DateAdd(DateInterval.Day, 2, tmpDeadlineDate)
            .Item("status") = "P"
            .Item("time_limit") = 0
            .Item("distance") = 0
            .Item("message_id") = 9
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '⑥	在定时任务表匹配workflow_id,project_id,task_id=RefundDebtInfo定时任务；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('RefundDebtInfo','RefundDebtInfoMsg')}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''异常处理  
        'If dsTempTimingTask.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsTempTimingTask.Tables(0))
        '    Throw wfErr
        'End If


        '⑦	将登记项目代偿提示任务的开始时间置为还款截止日期+6个月；将任务状态置为“P”；
        count = dsTempTimingTask.Tables(0).Rows.Count
        If count > 0 Then
            For j = 0 To count - 1
                dsTempTimingTask.Tables(0).Rows(j).Item("start_time") = DateAdd(DateInterval.Month, 6, tmpDeadlineDate)
                dsTempTimingTask.Tables(0).Rows(j).Item("status") = "P"
            Next
        End If
        'dsTempTimingTask.Tables(0).Rows(0).Item("start_time") = DateAdd(DateInterval.Month, 6, tmpDeadlineDate)
        'dsTempTimingTask.Tables(0).Rows(0).Item("status") = "P"

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        ''在定时任务表匹配workflow_id,project_id,task_id=RecordProjectTraceInfo登记保后检查记录表定时任务；
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('RecordProjectTraceInfo')}"
        'dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''将登记保后检查记录表的开始时间置为当前时间+60天；将任务状态置为“P”；
        'count = dsTempTimingTask.Tables(0).Rows.Count
        'If count > 0 Then
        '    For j = 0 To count - 1
        '        dsTempTimingTask.Tables(0).Rows(j).Item("start_time") = DateAdd(DateInterval.Day, 60, Now)
        '        dsTempTimingTask.Tables(0).Rows(j).Item("status") = "P"
        '    Next
        'End If


    End Function
End Class
