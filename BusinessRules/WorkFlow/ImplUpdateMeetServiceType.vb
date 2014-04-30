Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplUpdateMeetServiceType
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义评审费支出金额、收入总额
    Private TrialFeePayout, TotalTrialFeeIncome As Single

    Private WfProjectTask As WfProjectTask
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private project As project

    Private WorkflowType As WorkflowType
    Private workflow As WorkFlow

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        project = New Project(conn, ts)

        WorkflowType = New WorkflowType(conn, ts)
        workflow = New WorkFlow(conn, ts)

    End Sub


    '检查业务品种是否有变化
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获取项目阶段
        Dim strSql As String
        Dim tmpTaskPhase As String
        Dim dsTempProject, dsTask, dsTemp As DataSet
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempProject = project.GetProjectInfo(strSql)

        '异常处理  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        tmpTaskPhase = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("phase")), "", dsTempProject.Tables(0).Rows(0).Item("phase"))

        '根据业务品种和项目阶段获取模版ID
        strSql = "{service_type=" & "'" & workFlowID & "'" & " and isnull(phase,'')=" & "'" & tmpTaskPhase & "'" & "}"
        Dim dsWorkflowType As DataSet = WorkflowType.GetWorkflowTypeInfo(strSql)

        '异常处理  
        If dsWorkflowType.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsWorkflowType.Tables(0))
            Throw wfErr
        End If

        Dim strWorkflow As String = dsWorkflowType.Tables(0).Rows(0).Item("workflow_id")

        '判断该模版的实例是否已存在
        strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id=" & "'" & strWorkflow & "'" & "}"
        dsTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '如果存在(未改变业务品种)如果业务品种未改变且登记签约的任务状态为‘F’，启动登记签约后的放款条件任务
        If dsTask.Tables(0).Rows.Count <> 0 Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EndMeetRecord'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EditBussinessProcess'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            ''登记签约的任务状态为‘F’，启动登记签约后的放款条件任务
            'strSql = "{project_code=" & "'" & projectID & "'" & "and task_id='RecordSignature' and task_status='F'}"
            'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            'If dsTemp.Tables(0).Rows.Count > 0 Then
            '    workflow.ReLoanApplication(projectID)
            'End If

        Else '"业务品种改变了"

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EndMeetRecord'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EditBussinessProcess'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        End If

    End Function

End Class
