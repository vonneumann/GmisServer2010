'��Ŀ�ⶳ
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


Public Class ImplUnFreezeProject
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTimingTask As WfProjectTimingTask
    Private workflow As WorkFlow

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        workflow = New WorkFlow(conn, ts)

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ��Ŀ����ǰ5λ�뱾��Ŀ����Ŀ����ǰ5λ��ͬ�� RefundRecord��Ǽǻ���֤��������״̬ΪP����Ŀ��PN����Ŀ������PN��
        Dim sProjectCode As String = Mid(projectID, 1, 5)
        Dim strSql As String = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RefundRecord','RecordRefundCertificate') and task_status='P'}"
        Dim dsTemp, dsProjectNum As DataSet
        dsProjectNum = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim iPN, i, j, k As Integer
        iPN = dsProjectNum.Tables(0).Rows.Count

        '���PN=0������Ŀ����ǰ5λ�뱾��Ŀ����Ŀ����ǰ5λ��ͬ���ǼǱ������¼����¼������ٻ����Ŀ���ۡ���˱�����ټ���¼��,���˱������¼������״̬Ϊ��P����������ΪF��
        'ֹͣ��������ʱ����
        If iPN = 0 Then
            strSql = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess','CheckProjectProcess','CheckProjectTraceInfo') and task_status='P'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            strSql = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess','CheckProjectProcess','CheckProjectTraceInfo') and status='P'}"
            dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("status") = "E"
            Next
            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        Else
            ''����ֻ������һ����Ŀ�ı�����ٻ����������Ŀ����ǰ5λ�뱾��Ŀ����Ŀ����ǰ5λ��ͬ���ǼǱ������¼����¼������ٻ����Ŀ���ۡ���˱�����ټ���¼������״̬Ϊ��P����������Ϊ��������
            'strSql = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess','CheckProjectProcess','CheckProjectTraceInfo') and task_status='P'}"
            'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            'Dim tmpProjectCode As String = dsTemp.Tables(0).Rows(0).Item("project_code") '��ȡ��һ����Ŀ����Ŀ����

            'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            '    If dsTemp.Tables(0).Rows(i).Item("project_code") <> tmpProjectCode Then
            '        dsTemp.Tables(0).Rows(i).Item("task_status") = ""
            '    End If
            'Next

            'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            '2005-3-15 yjf edit �޸���Ŀ�ⶳ���ܣ�ȷ����ҵ������һ����Ŀ�б�����ٻ(�ҵ��ĵ�һ����Ŀ)
            Dim tmpProjectCodeNumOne, tmpProjectCode, tmpTaskID As String
            tmpProjectCodeNumOne = dsProjectNum.Tables(0).Rows(0).Item("project_code") '��ȡ��һ����Ŀ����Ŀ����

            strSql = "{project_code='" & tmpProjectCodeNumOne & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess')}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                tmpTaskID = dsTemp.Tables(0).Rows(i).Item("task_id")
                workflow.StartupTask(workFlowID, tmpProjectCodeNumOne, tmpTaskID, "", "")
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            'ȷ��������Ŀ�ı�����ٶ��ر�
            For j = 0 To dsProjectNum.Tables(0).Rows.Count - 1
                tmpProjectCode = dsProjectNum.Tables(0).Rows(j).Item("project_code")
                strSql = "{project_code='" & tmpProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess')}"
                dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                For k = 0 To dsTemp.Tables(0).Rows.Count - 1
                    If dsTemp.Tables(0).Rows(k).Item("project_code") <> tmpProjectCodeNumOne Then
                        dsTemp.Tables(0).Rows(k).Item("task_status") = ""
                    End If
                Next
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        End If
    End Function
End Class
