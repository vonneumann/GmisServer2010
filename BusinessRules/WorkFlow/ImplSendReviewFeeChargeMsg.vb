Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendReviewFeeChargeMsg
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    ''������Ŀ��������
    'Private project As project

    ''���嶨ʱ�����������
    'Private TimingServer As TimingServer


    ''����ͨ�ò�ѯ��������
    'Private CommonQuery As CommonQuery


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        ''ʵ������Ŀ����
        'project = New Project(conn, ts)

        ''ʵ������ʱ�����������
        'TimingServer = New TimingServer(conn, ts, True, True)

        ''ʵ����ͨ�ò�ѯ����
        'CommonQuery = New CommonQuery(conn, ts)

    End Sub

    '������ȡ�����ʱ֪ͨ��Ŀ������ȡ�����
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        ''��ȡ������Ա
        'Dim strsql As String = "{ProjectCode=" & "'" & projectID & "'" & "}"
        'Dim dsProjectInfo As DataSet = CommonQuery.GetProjectInfoEx(strsql)

        ''�쳣����  
        'If dsProjectInfo.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
        '    Throw wfErr
        'End If

        'Dim tmpAttend As String = dsProjectInfo.Tables(0).Rows(0).Item("13")
        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpAttend, "20", "N")


        '2010-05-13 yjf add ������Ŀ�������δԤ�������Ԥ����Ϣ
        Dim strSql As String = "{project_code is null}"
        Dim WfProjectTimingTask As New WfProjectTimingTask(conn, ts)
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        Dim newRow As DataRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "ReviewFeeCharge"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "24"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 15, Now)
            .Item("status") = "P"
            .Item("time_limit") = 15
            .Item("distance") = 0
            .Item("message_id") = 31
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "ReviewFeeCharge"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "29"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 15, Now)
            .Item("status") = "P"
            .Item("time_limit") = 15
            .Item("distance") = 0
            .Item("message_id") = 31
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "ReviewFeeCharge"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "21"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 15, Now)
            .Item("status") = "P"
            .Item("time_limit") = 15
            .Item("distance") = 0
            .Item("message_id") = 31
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = workFlowID
            .Item("project_code") = projectID
            .Item("task_id") = "ReviewFeeCharge"
            .Item("workflow_id") = workFlowID
            .Item("role_id") = "02"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 15, Now)
            .Item("status") = "P"
            .Item("time_limit") = 15
            .Item("distance") = 0
            .Item("message_id") = 31
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)
    End Function

End Class
