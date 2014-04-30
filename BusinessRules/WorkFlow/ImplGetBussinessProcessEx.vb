Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'���޸�����������̡����޸ġ�ҵ��Ʒ�֡�

Public Class ImplGetBussinessProcessEx
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '������Ŀ��������
    Private project As Project

    '����������������
    Private ConfTrial As ConfTrial

    '���幤�������Ͷ�������
    Private WorkflowType As WorkflowType

    '���幤������������
    Private WfProjectTask As WfProjectTask
    Private WfProjectMessages As WfProjectMessages
    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTimingTask As WfProjectTimingTask
    Private WorkFlow As WorkFlow

    Private WfProjectTrack As WfProjectTrack

    '����ͨ�ò�ѯ��������
    Private CommonQuery As CommonQuery

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������Ŀ����
        project = New Project(conn, ts)

        'ʵ����������������
        ConfTrial = New ConfTrial(conn, ts)

        'ʵ�������������Ͷ�������
        WorkflowType = New WorkflowType(conn, ts)

        'ʵ��������������
        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectMessages = New WfProjectMessages(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        WorkFlow = New WorkFlow(conn, ts)

        WfProjectTrack = New WfProjectTrack(conn, ts)

        'ʵ����ͨ�ò�ѯ����
        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim j As Integer
        Dim dsTemp, dsProjectInfo, dsAttend As DataSet
        Dim tmpWorkflowID, tmpManagerA, tmpManagerB As String
        Dim strTaskID As String

        '��õ�ǰ�ġ�RecordReviewConclution�������workflow_id
        strTaskID = "RecordReviewConclusion"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & strTaskID & "'" & "}"
        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        tmpWorkflowID = dsTemp.Tables(0).Rows(0).Item("workflow_id")

        '  ��ɾ������Ŀ����ʼ��"99"��"15"���������ģ��ʵ��
        strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id not in (" & "'" & tmpWorkflowID & "'" & ",'99','15')}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)
        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next
        WfProjectTask.UpdateWfProjectTask(dsTemp)


        '��ȡ��Ŀ����A,B
        strSql = "{ProjectCode=" & "'" & projectID & "'" & "}"
        dsProjectInfo = CommonQuery.GetProjectInfoEx(strSql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("24")
        tmpManagerB = dsProjectInfo.Tables(0).Rows(0).Item("25")

        '����������
        'WorkFlow.CreateProcess(workFlowID, projectID, userID)
        CopyTemplate(workFlowID, projectID)

        '��	����Ŀ����A��B��Ϊ�յ�Ա����Ϊ��Ŀ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='24' and attend_person=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For j = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(j).Item("attend_person") = tmpManagerA
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='25' and attend_person=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For j = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(j).Item("attend_person") = tmpManagerB
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)


    End Function

    Private Function CopyTemplate(ByVal workFlowID As String, ByVal projectID As String)
        Dim j As Integer
        Dim dsTemp, dsTemplate As DataSet
        Dim newRow As DataRow

        Dim strSql As String

        ''��ȡ��Ŀ�׶�
        Dim tmpTaskPhase, tmpTaskStatus As String
        Dim dsTempProject As DataSet
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

        Dim strWorkflowID As String = dsWorkflowType.Tables(0).Rows(0).Item("workflow_id")
        Dim strWorkflow As String = "workflow_id=" & "'" & strWorkflowID & "'"


        '����ģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("task_template", strWorkflow)
        dsTemp = WfProjectTask.GetWfProjectTaskInfo("null")

        For j = 0 To dsTemplate.Tables(0).Rows.Count - 1


            newRow = dsTemp.Tables(0).NewRow()
            With newRow

                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(j).Item("task_id")
                .Item("sequence") = dsTemplate.Tables(0).Rows(j).Item("sequence")
                .Item("task_name") = dsTemplate.Tables(0).Rows(j).Item("task_name")
                .Item("task_type") = dsTemplate.Tables(0).Rows(j).Item("task_type")
                .Item("apply_tool") = dsTemplate.Tables(0).Rows(j).Item("apply_tool")
                .Item("parameters") = dsTemplate.Tables(0).Rows(j).Item("parameters")
                .Item("duration") = dsTemplate.Tables(0).Rows(j).Item("duration")
                .Item("merge_relation") = dsTemplate.Tables(0).Rows(j).Item("merge_relation")
                .Item("flow_tool") = dsTemplate.Tables(0).Rows(j).Item("flow_tool")
                .Item("create_person") = dsTemplate.Tables(0).Rows(j).Item("create_person")
                .Item("create_date") = dsTemplate.Tables(0).Rows(j).Item("create_date")
                .Item("project_phase") = dsTemplate.Tables(0).Rows(j).Item("phase")
                .Item("project_status") = dsTemplate.Tables(0).Rows(j).Item("status")
                .Item("hasMessage") = dsTemplate.Tables(0).Rows(j).Item("hasMessage")

            End With

            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTask.UpdateWfProjectTask(dsTemp)

        '��ɫģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("task_role_template", strWorkflow)
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo("null")

        '����ɫģ����ӵ���ɫ����
        For j = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(j).Item("task_id")
                .Item("role_id") = dsTemplate.Tables(0).Rows(j).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '3��ת������ģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("task_transfer_template", strWorkflow)
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")

        '��ת������ģ����ӵ�ת����������
        For j = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(j).Item("task_id")
                .Item("next_task") = dsTemplate.Tables(0).Rows(j).Item("next_task")
                .Item("transfer_condition") = dsTemplate.Tables(0).Rows(j).Item("transfer_condition")
                .Item("project_status") = dsTemplate.Tables(0).Rows(j).Item("status")
                .Item("isItem") = dsTemplate.Tables(0).Rows(j).Item("isItem")
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)


        '4����ʱ����ģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("timing_task_template", strWorkflow)
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")

        '������ģ����ӵ�����ģ��ʵ������
        For j = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(j).Item("task_id")
                .Item("role_id") = dsTemplate.Tables(0).Rows(j).Item("role_id")
                .Item("distance") = dsTemplate.Tables(0).Rows(j).Item("distance")
                .Item("start_time") = "1900-01-01"
                .Item("message_id") = dsTemplate.Tables(0).Rows(j).Item("message_id")
                .Item("type") = dsTemplate.Tables(0).Rows(j).Item("type")
                .Item("time_limit") = dsTemplate.Tables(0).Rows(j).Item("time_limit")
                .Item("parameter") = dsTemplate.Tables(0).Rows(j).Item("parameter")
                .Item("create_person") = dsTemplate.Tables(0).Rows(j).Item("create_person")
                .Item("create_date") = dsTemplate.Tables(0).Rows(j).Item("create_date")
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        'ɾ�����޸������������̡�
        'delModifyConclusion(projectID)
    End Function

    'ɾ�����޸������������̡�workflow_id:15
    Private Sub delModifyConclusion(ByVal projectId As String)
        Dim strSql As String
        Dim j As Integer
        strSql = "{project_code='" & projectID & "' and workflow_id='15'}"
        Dim dsTemp As DataSet



        'ɾ�������˱�
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        'ɾ��ת�Ʊ�
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        'ɾ����ʱ�����
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        'ɾ�����ٱ�
        dsTemp = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next

        WfProjectTrack.UpdateWfProjectTrack(dsTemp)


        'ɾ�������
        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)

        For j = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(j).Delete()
        Next

        WfProjectTask.UpdateWfProjectTask(dsTemp)
    End Sub
End Class
