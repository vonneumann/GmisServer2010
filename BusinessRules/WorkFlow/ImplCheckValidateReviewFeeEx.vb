Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplCheckValidateReviewFeeEx
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���������֧���������ܶ�
    Private TrialFeePayout As Integer = 0
    Private TotalTrialFeeIncome As Integer = 0

    '������Ŀ�����������
    Private attendee As ProjectTaskAttendee


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������Ŀ�������
        attendee = New ProjectTaskAttendee(conn, ts)

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '�������Ŀ�ķ�������ʩ��������Э��е�30%��������ȷ���շѱ�׼����
        Dim projectGuaranteeForm As New ProjectGuaranteeForm(conn, ts)
        Dim strSql As String
        strSql = "{project_code='" & projectID & "' and guarantee_form like '%����Э��%' and isnull(is_used,0)=1}"
        Dim tmpDataSet As DataSet
        tmpDataSet = projectGuaranteeForm.GetProjectGuaranteeForm(strSql)
        Dim hasHuZhuHui As Boolean = (tmpDataSet.Tables(0).Rows.Count > 0)
        tmpDataSet.Dispose()

      

        Dim hasDoneConfirmReviewFee As Boolean
        Dim payoutIsBigger As Boolean

        '��������Ŀ����Ҫ��ȡ����ѣ�������Ҫ�����Ƿ�Ҫ������ѡ�
        If hasHuZhuHui = False Then
            Dim dsAttendee As DataSet = attendee.GetProjectTaskAttendeeInfo("{project_code='" & projectID & "' AND task_id='BalanceReviewFee'}")
            If dsAttendee.Tables(0).Rows.Count > 0 Then
                hasDoneConfirmReviewFee = CBool(dsAttendee.Tables(0).Rows(0)("task_status") & "" = "F")
            End If
            dsAttendee.Dispose()

            '��������ѵ������ܶ�
            GetTotalTrialFeeIncome(projectID)

            payoutIsBigger = (TotalTrialFeeIncome < TrialFeePayout)
        End If

        Dim i As Integer
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckValidateReviewFee'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
            If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "SendCashlossReviewMsg" Then
                dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = IIf(hasHuZhuHui, ".F.", IIf(hasDoneConfirmReviewFee AndAlso payoutIsBigger, ".T.", ".F."))
            Else
                dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = IIf(hasHuZhuHui, ".T.", IIf(hasDoneConfirmReviewFee AndAlso payoutIsBigger, ".F.", ".T."))
            End If
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function



    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '��ȡ����Ŀ��������ѵļ�¼
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='31'" & " and item_code='001'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '��������ѵ������ܶ�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout += CDbl(IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout")))
            TotalTrialFeeIncome += CDbl(IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
        Next
    End Function
End Class
