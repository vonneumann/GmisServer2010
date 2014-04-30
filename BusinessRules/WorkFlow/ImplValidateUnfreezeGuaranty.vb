Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplValidateUnfreezeGuaranty
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private WorkFlow As WorkFlow

    Private TimingServer As TimingServer

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        WorkFlow = New WorkFlow(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '�򵵰�������Ա�ͷ��չ�����������������������Ϣ
        '��Ŀ���÷�������ʩ��,�����������������ʩ����,��������Ϣ;����������Ŀ��ֹ

        Dim tmpFileManager, tmpMinister As String
        'tmpFileManager = WorkFlow.getTaskActor("42")
        'tmpMinister = WorkFlow.getTaskActor("31")

        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpFileManager, "23", "N")
        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpMinister, "23", "N")

        '��ò�����ת������
        Dim strSql As String
        Dim i As Integer
        Dim isOppGuarantee As Boolean = isHaveOppGuarantee(projectID) '�Ƿ�����˷�������ʩ
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectFinishedRrport'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        'If isOppGuarantee Then
        '    '�򵵰�������Ա�ͷ��չ�����������������������Ϣ
        '    tmpFileManager = WorkFlow.getTaskActor("42")
        '    tmpMinister = WorkFlow.getTaskActor("31")
        '    TimingServer.AddMsg(workFlowID, projectID, taskID, tmpFileManager, "23", "N")
        '    TimingServer.AddMsg(workFlowID, projectID, taskID, tmpMinister, "23", "N")

        '    For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
        '        If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "VaildateUnfreezeGuaranty" Then
        '            dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
        '        ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
        '            dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
        '        End If
        '    Next
        'Else
        '    For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
        '        If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "VaildateUnfreezeGuaranty" Then
        '            dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
        '        ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
        '            dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
        '        End If
        '    Next
        'End If

        If isOppGuarantee Or getDepositFee(projectID) Then

            ''�򵵰�������Ա�ͷ��չ�����������������������Ϣ
            'tmpFileManager = WorkFlow.getTaskActor("42")
            'tmpMinister = WorkFlow.getTaskActor("31")
            'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpFileManager, "23", "N")
            'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpMinister, "23", "N")

            '�б�֤�����ͷ�
            If getDepositFee(projectID) Then
                For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                    If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateReturnDepositFee" Then
                        dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                    ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                        dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                    End If
                Next
            End If

            '�з�����������
            If isOppGuarantee Then
                For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                    If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "VaildateUnfreezeGuaranty" Then
                        dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                    ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                        dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                    End If
                Next
            End If

        Else '�ޱ�֤���ͷ����޷���������
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "VaildateUnfreezeGuaranty" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateUnfreezeDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    '�Ƿ�����˷�������ʩ
    Private Function isHaveOppGuarantee(ByVal projectID As String) As Boolean
        Dim strSql As String
        Dim i, count As Integer
        Dim dsTemp As DataSet
        Dim projectGuarantee As New Guaranty(conn, ts)

        strSql = "{project_code='" & projectID & "' and guaranty_type<>'2A' and status like  '��Ѻ��Ѻ'}"
        dsTemp = projectGuarantee.GetGuarantyInfo(strSql, "null")
        count = dsTemp.Tables("opposite_guarantee").Rows.Count

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function


    '���ȷ����ȡ��֤��Ľ��
    Private Function getDepositFee(ByVal ProjectID As String) As Boolean

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

        If DepositFee > 0.0 Then
            Return True
        Else
            Dim projectGuarantee As New Guaranty(conn, ts)

            strSql = "{project_code='" & ProjectID & "' and guaranty_type='2A' and status like  '��Ѻ��Ѻ'}"
            dsTemp = projectGuarantee.GetGuarantyInfo(strSql, "null")
            Dim count As Integer
            count = dsTemp.Tables("opposite_guarantee").Rows.Count

            If count > 0 Then
                Return True
            Else
                Return False
            End If

        End If

    End Function
End Class
