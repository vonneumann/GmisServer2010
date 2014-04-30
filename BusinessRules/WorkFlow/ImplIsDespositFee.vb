
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�Ƿ���ȡ��֤��

Public Class ImplIsDespositFee
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���屣֤��֧���������ܶ�
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
        '���㱣֤��������ܶ�
        GetTotalTrialFeeIncome(projectID)

        Dim strSql As String

        Dim i As Integer
        Dim dsTempTaskTrans, dsAttend As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateDepositFee'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        If TotalTrialFeeIncome < TrialFeePayout Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "DepositFeeCharge" Or dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next

        Else 'ȷ���շ�Ϊ:0����payout=income,�������շ�������

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "DepositFeeCharge" Or dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

            '�������շ�����,��Ѹ�������Ϊ.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('DepositFeeCharge','ValidateDepositFee')}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '��ȡ����Ŀ���ڱ�֤��ļ�¼(31,002)
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='34'" & " and item_code='009'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '���㱣֤��������ܶ�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout = CInt(TrialFeePayout + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout")))
            TotalTrialFeeIncome = CInt(TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
        Next

    End Function

End Class
