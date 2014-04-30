Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'ȷ�Ϻ�����
Public Class ImplIsRecommendAffrim
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '����ת�������������
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private CooperateOpinion As CooperateOpinion


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

        CooperateOpinion = New CooperateOpinion(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        Dim dsCooperate, dsTempTaskTrans, dsAttend As DataSet
        '��	��Cooperate-Organization-Opinion���ȡ��Ŀ�ĺ���������Cooperate-Organization��
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsCooperate = CooperateOpinion.GetCooperateOpinionInfo("null", strSql)

        '��	�����ȡ�ĺ���������Ϊ��
        If dsCooperate.Tables(1).Rows.Count = 0 Then
            '    ��IsRecommendAffrim��RecommendAffrim��ת��������Ϊ.F.;

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsRecommendAffrim' and next_task='RecommendAffrim'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            '    ��RecommendAffrim������״̬��Ϊ��F��;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecommendAffrim'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        Else
            '����()
            '    ��IsRecommendAffrim��RecommendAffrim��ת��������Ϊ.T.;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsRecommendAffrim' and next_task='RecommendAffrim'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '�쳣����  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            Dim tmpCooperate As String = dsCooperate.Tables(1).Rows(0).Item("cooperate_organization")
            '    ��Cooperate-Organization���ȡCooperate-Organization�����ר��ԱManager;
            strSql = "{cooperate_organization=" & "'" & tmpCooperate & "'" & "}"
            dsCooperate = CooperateOpinion.GetCooperateOpinionInfo(strSql, "null")
            'qxd modify
            'Dim tmpManager As String = dsCooperate.Tables(0).Rows(0).Item("manager")
            Dim tmpManager As String = IIf(dsCooperate.Tables(0).Rows(0).Item("manager") Is System.DBNull.Value, "", dsCooperate.Tables(0).Rows(0).Item("manager"))

            '    ��RecommendAffrim����Ĳ�������ΪManager;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecommendAffrim'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpManager
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        End If
    End Function
End Class
