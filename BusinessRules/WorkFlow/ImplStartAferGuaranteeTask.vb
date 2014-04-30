Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplStartAferGuaranteeTask
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
        Dim dsTempAttend As DataSet

        '������Ǽ�������'P'(��Ϊ����ǼǵĶ�ʱ������Ѿ���������ʼ,����Ҫ����StartupTask����)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundRecord'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        '���ǼǱ�����ټ�¼������'P'(��Ϊ�ǼǱ�����ټ�¼�Ķ�ʱ������Ѿ���������ʼ,����Ҫ����StartupTask����)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordProjectTraceInfo'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

    End Function
End Class
