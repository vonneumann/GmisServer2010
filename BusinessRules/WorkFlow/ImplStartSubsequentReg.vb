Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

''��������Ǽ�����
'�رյǼǱ������¼����˱������¼����Ŀ��������
'ֹͣ�ǼǱ������¼���Ǽǻ��ʱ����
Public Class ImplStartSubsequentReg
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���嶨ʱ����
    Private WfProjectTimingTask As WfProjectTimingTask

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTempTaskAttendee As DataSet
        Dim i As Integer

        ''��	���Ǽ�����׷����Ϣ����TID=OverdueTrailRecord�����н�ɫ������״̬��Ϊ��P����
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='OverdueTrailRecord'}"
        'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
        '    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        ''��	���ǼǴ�����Ϣ����TID=RefundDebtInfo�����н�ɫ������״̬��Ϊ��P����
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundDebtInfo'}"
        'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
        '    dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        'Next

        '1.�رյǼǱ������¼����˱������¼����Ŀ��������
        strSql = "{project_code=" & "'" & projectID & "'" & " and (task_id='RecordProjectTraceInfo' or task_id='CheckProjectProcess' or task_id='AppraiseProjectProcess')}"

        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        '2.ֹͣ�ǼǱ������¼���Ǽǻ��ʱ����
        Dim dsTemp As DataSet
        strSql = "{project_code='" & projectID & "' and (task_id='RecordProjectTraceInfo' or task_id='RefundRecord')}"
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

    End Function

End Class
