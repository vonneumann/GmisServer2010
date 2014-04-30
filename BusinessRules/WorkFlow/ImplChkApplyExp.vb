Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplChkApplyExp
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义工作记录对象引用
    Private WorkLog As WorkLog

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    '定义工作流对象引用
    Private WorkFlow As WorkFlow

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化工作记录对象
        WorkLog = New WorkLog(conn, ts)

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        '实例化定时任务对象
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        '实例化转移任务任务对象
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        '实例化工作流对象引用
        WorkFlow = New WorkFlow(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        ''获取该人该项目的保前调查记录
        Dim strSql As String
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='PreguaranteeActivity' and attend_person=" & "'" & userID & "'" & "}"
        'Dim dsTempWorkLog As DataSet = WorkLog.GetWorkLogInfo(strSql)
        Dim dsTempTaskAttendee, dsTempTaskTrans As DataSet
        '2005-08-24 yjf edit 撤销保前调研的限制
        '①	如果worklog表记录非空
        'If dsTempWorkLog.Tables(0).Rows.Count <> 0 Then

        '将登记保前活动记录活动（TID=PreguaranteeActivity）所有角色的任务状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='PreguaranteeActivityExp'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        '   所有角色的任务状态置为“F”；
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '调用 DeleteAlert（模板ID、工作流ID、PreguaranteeActivity）；
        DeleteAlert(workFlowID, projectID, "PreguaranteeActivity")

        '如果登记反担保物（TID=ApplyCapitialEvaluated）的任务状态置为“P”，[本项目不需资产评估]
        '将申分配评估师活动（TID=ApplyCapitialEvaluated）的角色的任务状态置为“F”；
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ApplyCapitialEvaluatedExp'" & "}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        If dsTempTaskAttendee.Tables(0).Rows.Count <> 0 Then
            ''异常处理  
            'If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            '    Throw wfErr
            'End If

            '如果登记反担保物（TID=ApplyCapitialEvaluated）的任务状态置为“P”
            Dim tmpStatus As String = IIf(IsDBNull(dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status"))
            If tmpStatus = "P" Then
                'Dim tmpTaskID As String = dsTempTaskAttendee.Tables(0).Rows(0).Item("task_id")
                'strSql = "{project_code=" & "'" & projectID & "'"   & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '将登记反担保物活动的任务状态置为“F”
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
                Next

                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

                WorkFlow.AACKMassage(workFlowID, projectID, "ApplyCapitialEvaluated", userID)

                '调用 DeleteAlert（模板ID、工作流ID、AssignValuator）；
                DeleteAlert(workFlowID, projectID, "AssignValuator")

                '将评估资产任务（TID=CapitialEvaluated）所有角色的任务状态置为“F”；
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluatedExp'}"
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
                Next

                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

                WorkFlow.AACKMassage(workFlowID, projectID, "CapitialEvaluated", userID)

                ''调用 DeleteAlert（模板ID、工作流ID、CapitialEvaluated）；
                'DeleteAlert(workFlowID, projectID, "CapitialEvaluated")
            Else
                '将评估资产任务（TID=CapitialEvaluated）到记录评审会结论的转移条件置为真.T.
                '将评估资产任务（TID=CapitialEvaluated）到登记反担保物的转移条件置为假.F.
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluatedExp' and next_task='RecordReviewConclusionExp'}"
                dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                ''异常处理  
                'If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                '    Throw wfErr
                'End If

                For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Next
                WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluatedExp' and next_task='ApplyCapitialEvaluatedExp'}"
                dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                ''异常处理  
                'If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                '    Throw wfErr
                'End If

                For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Next

                WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            End If
        End If


        '' 2007-12-17 yjf add 添加展期上会记录标记
        'strSql = "{project_code='" & projectID & "' order by trial_time desc}"
        'Dim confTrial As New ConfTrial(conn, ts)
        'Dim dsExpContrial As DataSet = confTrial.GetConfTrialInfo(strSql, "")
        'If dsExpContrial.Tables(0).Rows.Count <> 0 Then
        '    dsExpContrial.Tables(0).Rows(0).Item("is_exp") = 1
        'End If
        'confTrial.UpdateConfTrial(dsExpContrial)

        'Else

        '    '否则
        '    '提示“未登记保前调查记录，提交调研结论失败！”；
        '    '将员工ID当前完成的任务状态置为“P”； 

        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and attend_person=" & "'" & userID & "'" & "}"
        '    dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status") = "P"
        '    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '    '抛出“未登记保前调查记录，提交调研结论失败
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowPreguaranteeActivityErr()
        '    Throw wfErr


        'End If

    End Function

    '将定时任务表中的当前任务ID（模板ID、项目ID、任务ID）状态改为“E”
    Public Function DeleteAlert(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String)

        '①	在定时任务表将与参数匹配的提示任务状态置为“E”；
        Dim strSql As String
        Dim dsTempTimingTask As DataSet
        Dim i As Integer
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        '   把指定任务ID的提示任务状态置为“E”
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

    End Function

End Class
