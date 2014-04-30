Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'������Ŀ��������
Public Class ImplEndProjectAppraise
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

    '�ر�����:��Ŀ����,��¼��ǰ���л,�ǼǷ�������ʩ,�ύ�������,�ύ���н���
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ProjectAppraiseReport'" & "}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('ProjectAppraiseReport','PreguaranteeActivity','ApplyCapitialEvaluated','ProjectAttitude','SubmissionProbeResult')" & "}"
        Dim dsTempAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)
    End Function
End Class
