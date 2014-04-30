
'�ж��Ƿ���Ҫ��������������ȷ��
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsValidateFraternity
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private OppositeGuaranteeForm As OppositeGuaranteeForm
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

        OppositeGuaranteeForm = New OppositeGuaranteeForm(conn, ts)
        ProjectGuaranteeForm = New ProjectGuaranteeForm(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        Dim dsFraternity, dsGuarantee, dsTempTaskTrans, dsAttend As DataSet

        '��	��Opposite_Guarantee_Form���л�ȡ������������ֵ䶨�壻
        strSql = "{form_code='06'}"
        dsFraternity = OppositeGuaranteeForm.GetOppositeGuaranteeForm(strSql)
        Dim tmpFraternity As String = Trim(dsFraternity.Tables(0).Rows(0).Item("name"))

        '��Project_Guarantee_Form�л�ȡ
        strSql = "{project_code=" & "'" & projectID & "'" & " and ltrim(rtrim(guarantee_form))='" & tmpFraternity & "' and is_used=1}"
        dsGuarantee = ProjectGuaranteeForm.GetProjectGuaranteeForm(strSql)

        '��	�����ȡ�Ļ��������Ϊ��
        If dsGuarantee.Tables(0).Rows.Count = 0 Then
            '    ��IsValidateFraternity��ValidateFraternity��ת��������Ϊ.F.;

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateFraternity' and next_task='ValidateFraternity'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            '    ��ValidateFraternity������״̬��Ϊ��F��;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateFraternity'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        Else
            '����()
            '    ��IsValidateFraternity��ValidateFraternity��ת��������Ϊ.T.;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateFraternity' and next_task='ValidateFraternity'}"
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
