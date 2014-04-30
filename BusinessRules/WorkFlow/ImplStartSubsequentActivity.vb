Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�������������
Public Class ImplStartSubsequentActivity
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

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


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTempTaskAttendee As DataSet
        Dim i As Integer

        '��	���ǼǱ�����ټ�¼����TID=RecordProjectProcess�����н�ɫ������״̬��Ϊ��P����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '��	������Ǽ�����TID=RefundRecord�����н�ɫ������״̬��Ϊ��P����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundRecord'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '��	���Ǽǻ���֤��������TID=RecordRefundCertificate�����н�ɫ������״̬��Ϊ��P����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordRefundCertificate'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '4 ��������Ŀ��չ����TID=AppraiseProjectProcess�����н�ɫ������״̬��Ϊ��P����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='AppraiseProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        '5 ����¼�������¼����TID=RecordProjectTraceInfo�����н�ɫ������״̬��Ϊ��P����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordProjectTraceInfo'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "P"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
    End Function

End Class
