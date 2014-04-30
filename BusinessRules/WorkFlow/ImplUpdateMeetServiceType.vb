Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplUpdateMeetServiceType
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���������֧���������ܶ�
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


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        project = New Project(conn, ts)

        WorkflowType = New WorkflowType(conn, ts)
        workflow = New WorkFlow(conn, ts)

    End Sub


    '���ҵ��Ʒ���Ƿ��б仯
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ��Ŀ�׶�
        Dim strSql As String
        Dim tmpTaskPhase As String
        Dim dsTempProject, dsTask, dsTemp As DataSet
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempProject = project.GetProjectInfo(strSql)

        '�쳣����  
        If dsTempProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempProject.Tables(0))
            Throw wfErr
        End If

        tmpTaskPhase = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("phase")), "", dsTempProject.Tables(0).Rows(0).Item("phase"))

        '����ҵ��Ʒ�ֺ���Ŀ�׶λ�ȡģ��ID
        strSql = "{service_type=" & "'" & workFlowID & "'" & " and isnull(phase,'')=" & "'" & tmpTaskPhase & "'" & "}"
        Dim dsWorkflowType As DataSet = WorkflowType.GetWorkflowTypeInfo(strSql)

        '�쳣����  
        If dsWorkflowType.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsWorkflowType.Tables(0))
            Throw wfErr
        End If

        Dim strWorkflow As String = dsWorkflowType.Tables(0).Rows(0).Item("workflow_id")

        '�жϸ�ģ���ʵ���Ƿ��Ѵ���
        strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id=" & "'" & strWorkflow & "'" & "}"
        dsTask = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�������(δ�ı�ҵ��Ʒ��)���ҵ��Ʒ��δ�ı��ҵǼ�ǩԼ������״̬Ϊ��F���������Ǽ�ǩԼ��ķſ���������
        If dsTask.Tables(0).Rows.Count <> 0 Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EndMeetRecord'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EditBussinessProcess'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            ''�Ǽ�ǩԼ������״̬Ϊ��F���������Ǽ�ǩԼ��ķſ���������
            'strSql = "{project_code=" & "'" & projectID & "'" & "and task_id='RecordSignature' and task_status='F'}"
            'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            'If dsTemp.Tables(0).Rows.Count > 0 Then
            '    workflow.ReLoanApplication(projectID)
            'End If

        Else '"ҵ��Ʒ�ָı���"

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EndMeetRecord'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsEditBussiness' and next_task='EditBussinessProcess'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
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
