'�ж��Ѿ��������
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsFinishedReturn
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private WfProjectTimingTask As WfProjectTimingTask

    Private ProjectAccountDetail As ProjectAccountDetail

    '����֧���������ܶ�
    Private TrialFeePayout As Single = 0
    Private TotalTrialFeeIncome As Single = 0

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ����ת���������
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        ProjectAccountDetail = New ProjectAccountDetail(conn, ts)

        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer

        '�ж���Ŀ�Ƿ�����ʧ
        strSql = "{project_code='" & projectID & "' and item_type='31' and item_code='004' and type='��ʧ'}"
        Dim dsProjectAccountDetail As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim iCount As Integer = dsProjectAccountDetail.Tables(0).Rows.Count

        Dim dsTempTaskTrans As DataSet
        strSql = "{project_code=" & "'" & ProjectID & "'" & " and task_id='RefundRecord'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


        '�������δ��ϲ���û����ʧ
        If TrialFeePayout <> TotalTrialFeeIncome And iCount <> 0 Then

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RefundRecord" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else

            '����

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RefundRecord" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        '������ǼǶ�ʱ��ʾ��Ϊ"E"
        strSql = "{project_code='" & projectID & "' and task_id='RefundRecord'}"
        Dim dsTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTimingTask.Tables(0).Rows.Count - 1
            dsTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

    End Function

    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '��ȡ����Ŀ��������ѵļ�¼
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code='" & ProjectID & "' item_type='31' and item_code='004' and type='����'}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '��������ѵ������ܶ�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout = TrialFeePayout + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout"))
            TotalTrialFeeIncome = TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income"))
        Next
    End Function
End Class
