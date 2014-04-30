Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplChkApply
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���幤����¼��������
    Private WorkLog As WorkLog

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���嶨ʱ�����������
    Private WfProjectTimingTask As WfProjectTimingTask

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    '���幤������������
    Private WorkFlow As WorkFlow

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ����������¼����
        WorkLog = New WorkLog(conn, ts)

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        'ʵ������ʱ�������
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        'ʵ����ת�������������
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        'ʵ������������������
        WorkFlow = New WorkFlow(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        ''��ȡ���˸���Ŀ�ı�ǰ�����¼
        Dim strSql As String
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='PreguaranteeActivity' and attend_person=" & "'" & userID & "'" & "}"
        'Dim dsTempWorkLog As DataSet = WorkLog.GetWorkLogInfo(strSql)
        Dim dsTempTaskAttendee, dsTempTaskTrans As DataSet
        '2005-08-24 yjf edit ������ǰ���е�����
        '��	���worklog���¼�ǿ�
        'If dsTempWorkLog.Tables(0).Rows.Count <> 0 Then

        '���ǼǱ�ǰ���¼���TID=PreguaranteeActivity�����н�ɫ������״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='PreguaranteeActivity'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        '   ���н�ɫ������״̬��Ϊ��F����
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '���� DeleteAlert��ģ��ID��������ID��PreguaranteeActivity����
        DeleteAlert(workFlowID, projectID, "PreguaranteeActivity")

        '����ǼǷ������TID=ApplyCapitialEvaluated��������״̬��Ϊ��P����[����Ŀ�����ʲ�����]
        '�����������ʦ���TID=ApplyCapitialEvaluated���Ľ�ɫ������״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ApplyCapitialEvaluated'" & "}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        If dsTempTaskAttendee.Tables(0).Rows.Count <> 0 Then
            ''�쳣����  
            'If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            '    Dim wfErr As New WorkFlowErr()
            '    wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            '    Throw wfErr
            'End If

            '����ǼǷ������TID=ApplyCapitialEvaluated��������״̬��Ϊ��P��
            Dim tmpStatus As String = IIf(IsDBNull(dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status")), "", dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status"))
            If tmpStatus = "P" Then
                'Dim tmpTaskID As String = dsTempTaskAttendee.Tables(0).Rows(0).Item("task_id")
                'strSql = "{project_code=" & "'" & projectID & "'"   & " and task_id=" & "'" & tmpTaskID & "'" & "}"
                'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '���ǼǷ�������������״̬��Ϊ��F��
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
                Next

                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

                WorkFlow.AACKMassage(workFlowID, projectID, "ApplyCapitialEvaluated", userID)

                '���� DeleteAlert��ģ��ID��������ID��AssignValuator����
                DeleteAlert(workFlowID, projectID, "AssignValuator")

                '�������ʲ�����TID=CapitialEvaluated�����н�ɫ������״̬��Ϊ��F����
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluated'}"
                dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
                Next

                WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

                WorkFlow.AACKMassage(workFlowID, projectID, "CapitialEvaluated", userID)

                ''���� DeleteAlert��ģ��ID��������ID��CapitialEvaluated����
                'DeleteAlert(workFlowID, projectID, "CapitialEvaluated")
            Else
                '�������ʲ�����TID=CapitialEvaluated������¼�������۵�ת��������Ϊ��.T.
                '�������ʲ�����TID=CapitialEvaluated�����ǼǷ��������ת��������Ϊ��.F.
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluated' and next_task='RecordReviewConclusion'}"
                dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                ''�쳣����  
                'If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                '    Dim wfErr As New WorkFlowErr()
                '    wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                '    Throw wfErr
                'End If

                For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Next
                WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluated' and next_task='ApplyCapitialEvaluated'}"
                dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

                ''�쳣����  
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

        'Else

        '    '����
        '    '��ʾ��δ�ǼǱ�ǰ�����¼���ύ���н���ʧ�ܣ�����
        '    '��Ա��ID��ǰ��ɵ�����״̬��Ϊ��P���� 

        '    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & " and attend_person=" & "'" & userID & "'" & "}"
        '    dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    dsTempTaskAttendee.Tables(0).Rows(0).Item("task_status") = "P"
        '    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '    '�׳���δ�ǼǱ�ǰ�����¼���ύ���н���ʧ��
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowPreguaranteeActivityErr()
        '    Throw wfErr


        'End If

    End Function

    '����ʱ������еĵ�ǰ����ID��ģ��ID����ĿID������ID��״̬��Ϊ��E��
    Public Function DeleteAlert(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String)

        '��	�ڶ�ʱ����������ƥ�����ʾ����״̬��Ϊ��E����
        Dim strSql As String
        Dim dsTempTimingTask As DataSet
        Dim i As Integer
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        '   ��ָ������ID����ʾ����״̬��Ϊ��E��
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

    End Function

End Class
