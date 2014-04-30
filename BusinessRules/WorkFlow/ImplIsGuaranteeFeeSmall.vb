
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�Ƿ���ȡ������

Public Class ImplIsGuaranteeFeeSmall
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���嵣����֧���������ܶ�
    Private TrialFeePayout As Integer = 0
    Private TotalTrialFeeIncome As Integer = 0

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '���㵣���ѵ������ܶ�
        GetTotalTrialFeeIncome(projectID)

        Dim strSql As String

        Dim i As Integer
        Dim dsTempTaskTrans, dsAttend As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateServiceFee'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        If GetTotalTrialFeeIncome(projectID) = True Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "SendGuaranteeChargeMsg" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next


        Else 'ȷ���շ�Ϊ:0����payout=income,�������շ�������

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "SendGuaranteeChargeMsg" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

            '�������շ�����,��Ѹ�������Ϊ.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ServiceFeeCharge'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String) As Boolean

        '��ȡ����Ŀ����С������ѵļ�¼(34,010)
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "' order by trial_times desc}"
        Dim objConfTrial As New ConfTrial(conn, ts)
        Dim dsTemp As DataSet = objConfTrial.GetConfTrialInfo(strSql, "")

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        If IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("service_rate")), 0, dsTemp.Tables(0).Rows(0).Item("service_rate")) > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

End Class
