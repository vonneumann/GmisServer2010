'��Ŀ����
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplFreezeProject
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

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

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ��Ŀ����ǰ5λ�뱾��Ŀ����Ŀ����ǰ5λ��ͬ��RefundRecord��Ǽǻ���֤��������״̬ΪP����Ŀ��PN��
        Dim sProjectCode As String = Mid(projectID, 1, 5)
        Dim strSql As String = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RefundRecord','RecordRefundCertificate') and task_status='P'}"
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim iPN, i As Integer
        iPN = dsTemp.Tables(0).Rows.Count

        '���PN��0����Tickling��ShareTrack����ת��������Ϊ�棬Tickling��Trackת�������ü٣������������ڶ�ʱ����ShareTrack���������䶳�����񣩣�
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='Ticking'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        If iPN > 0 Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ShareTrack" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ShareTrack" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

End Class
