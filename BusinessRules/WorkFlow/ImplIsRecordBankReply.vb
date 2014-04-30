Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsRecordBankReply
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
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

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)



    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        Dim dsTemp As DataSet
        '�ж�RecordBankReply�Ǽ����лָ������Ƿ�������
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordBankReply'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim tmpStatus As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("task_status")), "", dsTemp.Tables(0).Rows(0).Item("task_status"))

        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateReviewConclusion'}"
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


        '���������
        If tmpStatus = "F" And isAllowLaon(projectID) Then
            '          ��ValidateReviewConclusion��NewReviewConclusion��ת��������Ϊ.T.
            '          ��ValidateReviewConclusion�����������ת��������Ϊ.F.
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                If dsTemp.Tables(0).Rows(i).Item("next_task") = "NewReviewConclusion" Then
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        Else
            '����
            '          ��ValidateReviewConclusion��NewReviewConclusion��ת��������Ϊ.F.
            '          ��ValidateReviewConclusion�����������ת��������Ϊ.T.
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                If dsTemp.Tables(0).Rows(i).Item("next_task") = "NewReviewConclusion" Then
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)
    End Function

    'qxd add 2004-9-24
    '���Intent_letter�����лظ�����Ƿ���һ����¼Ϊ��ͬ�⡱��ע��һ����Ŀ���ֻ��һ����¼Ϊͬ�⣩��ͬ��������ſ�
    Private Function isAllowLaon(ByVal projectID As String) As Boolean
        Dim strSql As String
        Dim intentLetter As IntentLetter = New IntentLetter(conn, ts)
        Dim ds As DataSet

        strSql = "{project_code='" & projectID & "' and bank_reply='ͬ��'}"
        ds = intentLetter.GetIntentLetterInfo(strSql)
        If Not ds Is Nothing Then
            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
            Return False
        End If
    End Function
End Class
