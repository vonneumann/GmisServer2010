Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�жϷ��������Ƿ���Ҫ��������ʦ
Public Class ImplIsAssignValuator_Update
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    '���巴�������������
    Private Guaranty As Guaranty

    '������Ŀ�ֵ��������
    Private Item As Item

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

        'ʵ���������������
        Guaranty = New Guaranty(conn, ts)

        'ʵ������Ŀ�����ֵ����
        Item = New Item(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim i, j As Integer
        Dim strSql As String
        Dim dsGuaranty, dsItem, dsTempTaskTrans As DataSet
        Dim tmpItemCode As String
        Dim isAssignValuator As Boolean

        '�жϸ���Ŀ�ķ��������Ƿ񶼲���Ҫ��������ʦ
        strSql = "{project_code=" & "'" & projectID & "'" & " and evaluate_person is null}"
        dsGuaranty = Guaranty.GetGuarantyInfo(strSql, "null")

        For i = 0 To dsGuaranty.Tables(0).Rows.Count - 1
            tmpItemCode = dsGuaranty.Tables(0).Rows(i).Item("guaranty_type")
            dsItem = Item.GetItemType(tmpItemCode)

            '�쳣����  
            If dsItem.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsItem.Tables(0))
                Throw wfErr
            End If

            If IIf(IsDBNull(dsItem.Tables(0).Rows(0).Item("additional_remark")), 0, dsItem.Tables(0).Rows(0).Item("additional_remark")) = 1 Then

                isAssignValuator = True

                Exit For

            End If
        Next


        If isAssignValuator = False Then


            '���ж��Ƿ��ʲ�����������TID=CHKneedEvaluated���������ʲ�����ʦ��ת��������Ϊ��.F.
            '���ж��Ƿ��ʲ�����������TID=CHKneedEvaluated�����ǼǷ��������ת��������Ϊ��.T.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated_Update' and next_task='AssignValuator_Update'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated_Update' and next_task='EndUpdateGuarantee'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        Else
            '����

            '���ж��Ƿ��ʲ�����������TID=CHKneedEvaluated���������ʲ�����ʦ��ת��������Ϊ��.T.
            '���ж��Ƿ��ʲ�����������TID=CHKneedEvaluated�����ǼǷ��������ת��������Ϊ��.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated_Update' and next_task='AssignValuator_Update'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKneedEvaluated_Update' and next_task='EndUpdateGuarantee'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
        End If


    End Function
End Class