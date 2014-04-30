Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplInitialCapitialPath
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

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ����ת�������������
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim dsTempTaskTrans As DataSet
        '�������ʲ�����TID=CapitialEvaluated������¼�������۵�ת��������Ϊ��.F.
        '�������ʲ�����TID=CapitialEvaluated�����ǼǷ��������ת��������Ϊ��.T.
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluated' and next_task='RecordReviewConclusion'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '�쳣����  
        If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
            Throw wfErr
        End If

        dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluated' and next_task='ApplyCapitialEvaluated'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '�쳣����  
        If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
            Throw wfErr
        End If

        dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        '2009-09-16 yjf add 
        '���������������״̬��Ϊ�գ����������ϻ�����õ������п�����Ϊ��״̬Ϊ��ɣ���������
        strSql = "{project_code='" & projectID & "' and task_id='ReviewMeetingPlan'}"
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

    End Function
End Class
