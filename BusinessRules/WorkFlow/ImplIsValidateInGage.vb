Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsValidateInGage
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private ProjectGuaranteeForm As ProjectGuaranteeForm


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

        ProjectGuaranteeForm = New ProjectGuaranteeForm(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim i As Integer
        Dim strSql As String
        '��Project_Guarantee_Form���л�ȡ��Ѻ���¼
        strSql = "{project_code=" & "'" & projectID & "'" & " and guarantee_form='��Ѻ' and is_used=1}"
        Dim dsGuarantee, dsTempTaskTrans, dsAttend As DataSet
        dsGuarantee = ProjectGuaranteeForm.GetProjectGuaranteeForm(strSql)

        '�����Ѻ���¼Ϊ��
        If dsGuarantee.Tables(0).Rows.Count = 0 Then
            ' ��IsValidateInGage��ValidateInGageת������.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateInGage' and next_task ='ValidateInGage'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            ' ��ValidateInGage��״̬��ΪF
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id ='ValidateInGage'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        Else
            '����
            ' ��ImplIsValidateInGage��ValidateInGageת������.T.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateInGage' and next_task='ValidateInGage'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
        End If
    End Function
End Class
