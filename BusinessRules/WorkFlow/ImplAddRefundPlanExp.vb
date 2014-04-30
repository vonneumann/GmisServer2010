Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplAddRefundPlanExp
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '������Ŀ�ʻ���ϸ��������
    Private ProjectAccountDetail As ProjectAccountDetail

    '���嶨ʱ�����������
    Private WfProjectTimingTask As WfProjectTimingTask

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private TracePlan As TracePlan


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������Ŀ�ʻ���ϸ����
        ProjectAccountDetail = New ProjectAccountDetail(conn, ts)

        'ʵ������ʱ�������
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        'ʵ���������˶���
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

        '��	��Project-account-detial���ȡitem-code=002��item-type=34�����л��ʼʱ��date��
        strSql = "SELECT DISTINCT project_code, date" & _
                 " FROM dbo.project_account_detail" & _
                 " WHERE item_code='002' and item_type='34'" & _
                 " and project_code = " & "'" & projectID & "'"
        dsTempAccountDetail = CommonQuery.GetCommonQueryInfo(strSql)

        '��	Ϊÿ������ƻ��ڶ�ʱ�������뻹����ʾ����

        '��ɾ��ԭ�еĻ�����ʾ
        strSql = "{project_code='" & projectID & "' and task_id='RefundRecord'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        'ģ��ID= workflow_id��
        '��ĿID= project_id��
        '����ID= task_id��
        '��ɫID=24��
        '        ���� = "P"
        '��ʼʱ��=date
        '        ��� = 0
        '����ʱ����״̬��Ϊ��P����
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

        ''��ӱ�����ټ�¼�Ķ�ʱ��ʾ����
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

        '��	�ڶ�ʱ�����ƥ��workflow_id,project_id,task_id=OverdueRecord��ʱ����
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('OverdueRecord','OverdueRecordMsg')}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('OverdueRecord') and type='T'}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''�쳣����  
        'If dsTempTimingTask.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsTempTimingTask.Tables(0))
        '    Throw wfErr
        'End If


        '��	��ȡ���������ֹ���ڣ�
        'strSql = "{project_code=" & "'" & projectID & "'" & "}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and  isnull(end_date,'')<>''}"
        Dim LoanNotice As New LoanNotice(conn, ts)
        Dim dsTempLoanNotice As DataSet = LoanNotice.GetLoanNoticeInfo(strSql)

        '�쳣����  
        If dsTempLoanNotice.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempLoanNotice.Tables(0))
            Throw wfErr
        End If

        tmpDeadlineDate = dsTempLoanNotice.Tables(0).Rows(0).Item("end_date")
        '����ʱ��Ϊ��������+1��
        tmpDeadlineDate = DateAdd(DateInterval.Day, 1, tmpDeadlineDate)

        '��	���Ǽ���Ŀ������ʾ����Ŀ�ʼʱ����Ϊ�����ֹ���ڣ�������״̬��Ϊ��P����
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

        '��Ŀ�ڻ�����յĵڶ���û���ύ����֤��������ʱ,����Ŀ��������Ŀ������Ϣ
        '���������������û���ύ����֤��������,�����Ρ��������ܡ������������������û��ڶ�ʱ����ģ�涨���ɫΪ׼��������Ŀ������Ϣ

        '��ɾ��ԭ�е�������ʾ
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

        '��	�ڶ�ʱ�����ƥ��workflow_id,project_id,task_id=RefundDebtInfo��ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('RefundDebtInfo','RefundDebtInfoMsg')}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''�쳣����  
        'If dsTempTimingTask.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsTempTimingTask.Tables(0))
        '    Throw wfErr
        'End If


        '��	���Ǽ���Ŀ������ʾ����Ŀ�ʼʱ����Ϊ�����ֹ����+6���£�������״̬��Ϊ��P����
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

        ''�ڶ�ʱ�����ƥ��workflow_id,project_id,task_id=RecordProjectTraceInfo�ǼǱ������¼��ʱ����
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('RecordProjectTraceInfo')}"
        'dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        ''���ǼǱ������¼��Ŀ�ʼʱ����Ϊ��ǰʱ��+60�죻������״̬��Ϊ��P����
        'count = dsTempTimingTask.Tables(0).Rows.Count
        'If count > 0 Then
        '    For j = 0 To count - 1
        '        dsTempTimingTask.Tables(0).Rows(j).Item("start_time") = DateAdd(DateInterval.Day, 60, Now)
        '        dsTempTimingTask.Tables(0).Rows(j).Item("status") = "P"
        '    Next
        'End If


    End Function
End Class
