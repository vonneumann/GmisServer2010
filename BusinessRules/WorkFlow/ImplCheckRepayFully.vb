Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplCheckRepayFully
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ø���Ŀ�ķſ��ܶ�
        Dim TotalLoanIncome As Double
        TotalLoanIncome = getTotalLoan(projectID) * 10000.0 '��λ��Ԫת��Ϊ��Ԫ

        '�жϻ�����Ƿ���ڵ��ڷſ��ǣ�����ת���Ǽǻ���֤���飻������ת���Ǽǻ���
        Dim strSql As String
        Dim i As Integer

        '��ò����㻹����ܶ�
        strSql = "{project_code=" & "'" & projectID & "'" & " and item_type='34' and item_code='001'}"
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim TotalTrialFeeIncome As Double '�����ܶ�

        '�����ܶ�
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalTrialFeeIncome = CDbl(TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
            Next
        End If

        '��ò�����ת������
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKRepayFully'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        If TotalTrialFeeIncome >= TotalLoanIncome Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RecordRefundCertificate" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RecordRefundCertificate" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    '
    Private Function getTotalLoan(ByVal ProjectID As String) As Double

        Dim TotalLoanIncome As Double = 0.0

        '��ȡ����Ŀ���ڷſ��ļ�¼
        Dim ProjectLoanNotice As New LoanNotice(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'}"
        Dim dsTemp As DataSet = ProjectLoanNotice.GetLoanNoticeInfo(strSql)
        Dim i As Integer

        '����ſ��ܶ�
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalLoanIncome = CDbl(TotalLoanIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("sum")), 0, dsTemp.Tables(0).Rows(i).Item("sum")))
            Next
        End If

        Return TotalLoanIncome
    End Function
End Class
