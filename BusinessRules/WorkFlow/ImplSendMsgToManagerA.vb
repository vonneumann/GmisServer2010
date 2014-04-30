
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendMsgToManagerA
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���嶨ʱ�����������
    Private TimingServer As TimingServer

    '����ͨ�ò�ѯ��������
    Private CommonQuery As CommonQuery


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������ʱ�����������
        TimingServer = New TimingServer(conn, ts, True, True)

        'ʵ����ͨ�ò�ѯ����
        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    '������ȡ�����ʱ֪ͨ��Ŀ������ȡ�����
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '��ȡ��Ŀ����A,B
        Dim tmpManagerA As String
        Dim strsql As String = "select manager_A from queryProjectInfo where ProjectCode='" & projectID & "'"
        Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strsql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("manager_A")
        Dim dsOpp As DataSet
        Dim strAffirmEvaluateDate As String
        strsql = "select * from opposite_guarantee where project_code='" & projectID & "'and evaluate_person='" & userID & "'"
        dsOpp = CommonQuery.GetCommonQueryInfo(strsql)
        If dsOpp.Tables(0).Rows.Count > 0 Then
            strAffirmEvaluateDate = "ȷ����������Ϊ��" & dsOpp.Tables(0).Rows(0).Item("affirm_evaluate_date")
        End If
        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpManagerA, "28", "N")
        TimingServer.AddMsgContent(workFlowID, projectID, taskID, tmpManagerA, strAffirmEvaluateDate, "N")

    End Function



End Class
