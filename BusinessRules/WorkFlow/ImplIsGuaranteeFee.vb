
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�Ƿ���ȡ������

Public Class ImplIsGuaranteeFee
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
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateGuaranteeFee'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        If TotalTrialFeeIncome < TrialFeePayout Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "SendGuaranteeChargeMsg" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next

            '���յ�����
            Dim objProject As New Project(conn, ts)
            Dim dsTempProject As DataSet = objProject.GetProjectInfo("{project_code='" & projectID & "'}")
            Dim bIsFee As Boolean
            Dim refeeDate As DateTime
            If dsTempProject.Tables(0).Rows.Count <> 0 Then
                bIsFee = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("is_refee")), 0, dsTempProject.Tables(0).Rows(0).Item("is_refee"))
            End If

            If bIsFee Then

                refeeDate = IIf(IsDBNull(dsTempProject.Tables(0).Rows(0).Item("refee_date")), "1900-01-01", dsTempProject.Tables(0).Rows(0).Item("refee_date"))

                '��ɾ��ԭ�е���ʾ
                strSql = "{project_code='" & projectID & "' and task_id='ValidateGuaranteeFeeEx'}"

                Dim WfProjectTimingTask As New WfProjectTimingTask(conn, ts)
                Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
                For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
                    dsTempTimingTask.Tables(0).Rows(i).Delete()
                Next
                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

                Dim newRow As DataRow = dsTempTimingTask.Tables(0).NewRow
                With newRow
                    .Item("workflow_id") = workFlowID
                    .Item("project_code") = projectID
                    .Item("task_id") = "ValidateGuaranteeFeeEx"
                    .Item("role_id") = "43"
                    .Item("type") = "T"
                    .Item("start_time") = refeeDate
                    .Item("status") = "P"
                    .Item("time_limit") = 0
                    .Item("distance") = 0
                    .Item("message_id") = 30
                End With
                dsTempTimingTask.Tables(0).Rows.Add(newRow)

                WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)
            End If


        Else 'ȷ���շ�Ϊ:0����payout=income,�������շ�������

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "SendGuaranteeChargeMsg" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

            '�������շ�����,��Ѹ�������Ϊ.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='GuaranteeCharge'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '��ȡ����Ŀ���ڵ����ѵļ�¼(31,002)
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='31'" & " and item_code='002'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '���㵣���ѵ������ܶ�
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout = CInt(TrialFeePayout + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout")))
            TotalTrialFeeIncome = CInt(TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
        Next

    End Function

End Class
