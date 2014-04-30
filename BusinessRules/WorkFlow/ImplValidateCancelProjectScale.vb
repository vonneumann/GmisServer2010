Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplValidateCancelProjectScale
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private project As Project

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

        project = New Project(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String

        '��ȡ��Ŀ�Ƿ��ٴ�����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsTemp As DataSet = project.GetProjectInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim applySum As Single = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("apply_sum")), 0, dsTemp.Tables(0).Rows(0).Item("apply_sum"))

        'ValidateProjectScale�ӿ�ʵ�ֵ���������ڵ���1000��Ԫʱ����ValidateCancelProjectScale-IsReturnGuaranteeFee��Ϊ�٣�ValidateProjectScale-ValidateCancleProject��Ϊ�档
        If applySum >= 1000 Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateCancelProjectScale' and next_task='IsReturnGuaranteeFee'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateCancelProjectScale' and next_task='ValidateCancleProject'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        Else

            '����,��ValidateCancelProjectScale-IsReturnGuaranteeFee��Ϊ�棬ValidateCancelProjectScale-ValidateCancleProject��Ϊ�١�

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateCancelProjectScale' and next_task='IsReturnGuaranteeFee'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateCancelProjectScale' and next_task='ValidateCancleProject'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        End If

    End Function

End Class
