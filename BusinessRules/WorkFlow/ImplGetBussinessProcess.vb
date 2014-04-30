Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


Public Class ImplGetBussinessProcess
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '������Ŀ��������
    Private project As project

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
        Dim i As Integer
        Dim dsTemp, dsProjectInfo, dsAttend As DataSet
        Dim tmpWorkflowID, tmpManagerA, tmpManagerB As String

        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        tmpWorkflowID = dsTemp.Tables(0).Rows(0).Item("workflow_id")

        '  ��ɾ������Ŀ����ʼ��"99"���������ģ��ʵ��
        strSql = "{project_code=" & "'" & projectID & "'" & " and workflow_id not in (" & "'" & tmpWorkflowID & "'" & ",'99')}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next
        WfProjectTask.UpdateWfProjectTask(dsTemp)


        '��ȡ��Ŀ����A,B
        strSql = "select nowManagerA,nowManagerB from queryProjectInfo where projectCode='" & projectID & "'"
        dsProjectInfo = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("nowManagerA")
        tmpManagerB = dsProjectInfo.Tables(0).Rows(0).Item("nowManagerB")

        '��ȡ������¼Ա����Ϊ���ڵķ�����
        strSql = "select top 1 attend_person from project_task_attendee where project_code='" & projectID & "' and role_id='51'"
        Dim dsAttendee As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsAttendee.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsAttendee.Tables(0))
            Throw wfErr
        End If

        Dim tmpRecordPerson As String = dsAttendee.Tables(0).Rows(0).Item("attend_person")


        '2010-10-12 yjf edit ���ú��ڷ�������
        strSql = "select GuaranteeSum  from queryProjectInfo where projectCode='" & projectID & "'"
        dsProjectInfo = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        Dim dGuaranteeSum As Double = dsProjectInfo.Tables(0).Rows(0).Item("GuaranteeSum")

        'If tmpRecordPerson = "��ϼ" Then
        '    If dGuaranteeSum > 999.9 Then
        '        tmpRecordPerson = "Ф��"
        '    Else
        '        tmpRecordPerson = "��ʤ��"
        '    End If
        'ElseIf tmpRecordPerson = "Ф��" Then
        '    tmpRecordPerson = "����"
        'ElseIf tmpRecordPerson = "��һ��" Then
        '    tmpRecordPerson = "��һ��"
        'ElseIf tmpRecordPerson = "��ʤ��" Then
        '    tmpRecordPerson = "���Ƹ�"
        'End If

    



        '����������
        'WorkFlow.CreateProcess(workFlowID, projectID, userID)
        CopyTemplate(workFlowID, projectID)

        '��	����Ŀ����A��B��Ϊ�յ�Ա����Ϊ��Ŀ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='24' and attend_person=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpManagerA
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='25' and attend_person=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpManagerB
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        ''���÷�����
        'strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='33' and attend_person=''}"
        'dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsAttend.Tables(0).Rows.Count - 1
        '    dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpRecordPerson
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)


        '2011-5-20 YJF ADD 
        '���÷�����
        '��ȡ��Ŀ�������ڵĲ���
        strSql = "select dept_name from staff where staff_name='" & tmpManagerA & "'"
        dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
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


        '���ñ���Ŀ�ķ�����
        strSql = "{project_code='" & projectID & "' and role_id='33'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
            drTemp.Item("attend_person") = strPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)


    End Function

    Private Function CopyTemplate(ByVal workFlowID As String, ByVal projectID As String)
        Dim i As Integer
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

        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1


            newRow = dsTemp.Tables(0).NewRow()
            With newRow

                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
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

        '��ɫģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("task_role_template", strWorkflow)
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo("null")

        '����ɫģ����ӵ���ɫ����
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '3��ת������ģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("task_transfer_template", strWorkflow)
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo("null")

        '��ת������ģ����ӵ�ת����������
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("next_task") = dsTemplate.Tables(0).Rows(i).Item("next_task")
                .Item("transfer_condition") = dsTemplate.Tables(0).Rows(i).Item("transfer_condition")
                .Item("project_status") = dsTemplate.Tables(0).Rows(i).Item("status")
                .Item("isItem") = dsTemplate.Tables(0).Rows(i).Item("isItem")
            End With
            dsTemp.Tables(0).Rows.Add(newRow)
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)


        '4����ʱ����ģ��
        dsTemplate = WorkFlow.GetWfProjectTaskTemplateInfo("timing_task_template", strWorkflow)
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo("null")

        '������ģ����ӵ�����ģ��ʵ������
        For i = 0 To dsTemplate.Tables(0).Rows.Count - 1
            newRow = dsTemp.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = strWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
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

        'ɾ�����޸������������̡�
        'delModifyConclusion(projectID)
    End Function

    'ɾ�����޸������������̡�workflow_id:15
    Private Sub delModifyConclusion(ByVal projectId As String)
        Dim strSql As String
        Dim i As Integer
        strSql = "{project_code='" & projectID & "' and workflow_id='15'}"
        Dim dsTemp As DataSet



        'ɾ�������˱�
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        'ɾ��ת�Ʊ�
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        'ɾ����ʱ�����
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        'ɾ�����ٱ�
        dsTemp = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTrack.UpdateWfProjectTrack(dsTemp)


        'ɾ�������
        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTask.UpdateWfProjectTask(dsTemp)
    End Sub

End Class
