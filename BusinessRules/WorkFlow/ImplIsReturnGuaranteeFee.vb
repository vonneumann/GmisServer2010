'�ж��Ƿ����뵣����
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsReturnGuaranteeFee
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private ProjectAccountDetail As ProjectAccountDetail

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

        ProjectAccountDetail = New ProjectAccountDetail(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        '��ȡ��������Ϣ
        strSql = "{project_code='" & projectID & "' and item_type='31' and item_code='002' and income is not null}"
        Dim dsAccount, dsTemp As DataSet
        dsAccount = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer
        Dim sIncome As Single

        '���㵣������ȡ�ܶ�
        For i = 0 To dsAccount.Tables(0).Rows.Count - 1
            sIncome = sIncome + dsAccount.Tables(0).Rows(i).Item("income")
        Next

        '���������δ��
        ' ��IsReturnGuaranteeFee-ReturnGuaranteeFee��Ϊ�٣�IsReturnGuaranteeFee-SubmitCancelProArchives��Ϊ�档
        If sIncome = 0 Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='ReturnGuaranteeFee'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='SubmitCancelProArchives'}"
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

            '����,��IsReturnGuaranteeFee-ReturnGuaranteeFee��Ϊ�棬IsReturnGuaranteeFee-SubmitCancelProArchives��Ϊ�١�

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='ReturnGuaranteeFee'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='SubmitCancelProArchives'}"
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
