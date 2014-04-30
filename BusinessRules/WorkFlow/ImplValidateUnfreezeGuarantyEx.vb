
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplValidateUnfreezeGuarantyEx
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

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

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��Ŀ���÷�������ʩ��,�����������������ʩ����,��������Ϣ;����������Ŀ��ֹ

        '����ſ������
        Dim TrialFeePayout As Double = 0.0
        Dim TotalTrialFeeIncome As Double = 0.0
        Dim DepositFee As Double = 0.0

        TrialFeePayout = getTotalLoan(projectID)
        TotalTrialFeeIncome = getTotalTrialFeeIncome(projectID)
        DepositFee = getDepositFee(projectID)

        Dim strSql As String

        Dim i As Integer
        Dim dsTempTaskTrans, dsAttend As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectFinishedRrport'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '��������ɲ��ұ�֤������ȡ
        'If TotalTrialFeeIncome >= TrialFeePayout And DepositFee > 0.0 Then
        If DepositFee > 0.0 Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                'If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "VaildateUnfreezeGuaranty" Then
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateReturnDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateReturnDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    '���ȷ����ȡ��֤��Ľ��
    Private Function getDepositFee(ByVal ProjectID As String) As Double

        '��ȡ����Ŀ���ڱ�֤��ļ�¼(34,009)
        Dim DepositFee As Double = 0.0
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='34'" & " and item_code='009'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '���㱣֤��������ܶ�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            DepositFee = CDbl(DepositFee + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout")))
        Next

        Return DepositFee

    End Function

    '��ȡ����Ŀ���ڷſ��Ľ��
    Private Function getTotalLoan(ByVal ProjectID As String) As Double

        Dim TotalLoan As Double = 0.0

        '��ȡ����Ŀ���ڷſ��ļ�¼
        Dim ProjectLoanNotice As New LoanNotice(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'}"
        Dim dsTemp As DataSet = ProjectLoanNotice.GetLoanNoticeInfo(strSql)
        Dim i As Integer

        '����ſ��ܶ�
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalLoan = CDbl(TotalLoan + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("sum")), 0, dsTemp.Tables(0).Rows(i).Item("sum")))
            Next
        End If

        Return TotalLoan
    End Function

    '��ò����㻹����ܶ�
    Private Function getTotalTrialFeeIncome(ByVal ProjectID As String) As Double
        Dim strSql As String
        Dim i As Integer

        '��ò����㻹����ܶ�
        Dim TotalTrialFee As Double = 0.0
        strSql = "{project_code=" & "'" & ProjectID & "'" & " and item_type='34' and item_code='001'}"
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim TotalTrialFeeIncome As Double '�����ܶ�

        '�����ܶ�
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalTrialFee = CDbl(TotalTrialFee + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
            Next
        End If

        Return TotalTrialFee
    End Function
End Class
