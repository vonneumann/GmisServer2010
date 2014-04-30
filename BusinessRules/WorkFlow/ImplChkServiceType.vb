Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplChkServiceType
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义评审费支出金额、收入总额
    Private TrialFeePayout, TotalTrialFeeIncome As Single

    Private WfProjectTask As WfProjectTask
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private project As project

    Private WorkflowType As WorkflowType

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

        project = New Project(conn, ts)

        WorkflowType = New WorkflowType(conn, ts)

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

        Dim i As Integer
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsChangeBussiness'}"
        Dim dsTempTaskTrans As DataSet = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


        '如果存在
        If dsTask.Tables(0).Rows.Count <> 0 Then

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "GetBussinessProcess" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsChangeBussiness' and next_task='ValidateReviewConclusion'}"
            'dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            ''异常处理  
            'If dsTemp.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            '    Throw wfErr
            'End If

            'dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            'WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsChangeBussiness' and next_task='GetBussinessProcess'}"
            'dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            ''异常处理  
            'If dsTemp.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            '    Throw wfErr
            'End If

            'dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            'WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)
        Else

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "GetBussinessProcess" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next

            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsChangeBussiness' and next_task='ValidateReviewConclusion'}"
            'dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            ''异常处理  
            'If dsTemp.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            '    Throw wfErr
            'End If

            'dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            'WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsChangeBussiness' and next_task='GetBussinessProcess'}"
            'dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            ''异常处理  
            'If dsTemp.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            '    Throw wfErr
            'End If

            'dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            'WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        ''2011-5-20 YJF ADD 
        ''设置法务经理
        ''获取项目经理所在的部门

        ''获取项目经理A,B
        'Dim CommonQuery = New CommonQuery(conn, ts)
        'strSql = "select nowManagerA,nowManagerB from queryProjectInfo where projectCode='" & projectID & "'"
        'Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        ''异常处理  
        'If dsProjectInfo.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
        '    Throw wfErr
        'End If

        'Dim tmpManagerA As String = dsProjectInfo.Tables(0).Rows(0).Item("nowManagerA")


        'strSql = "select dept_name from staff where staff_name='" & tmpManagerA & "'"
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

        ''异常处理  
        'If dsTemp.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '    Throw wfErr
        'End If

        'Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        'strSql = "select staff_name from staff where  isnull(unchain_department_list,'') like '%" & strDeptName & "%'"
        'Dim dsTemp2 As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        'Dim strPerson As String
        'If dsTemp2.Tables(0).Rows.Count <> 0 Then
        '    strPerson = dsTemp2.Tables(0).Rows(0).Item("staff_name")
        'End If


        ''设置本项目的法务经理
        'Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
        'strSql = "{project_code='" & projectID & "' and role_id='33'}"
        'Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
        '    drTemp.Item("attend_person") = strPerson
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

    End Function
End Class
