Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


'���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
Public Class ImplSendMakeContractMsg
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '�����ɫ�û���������
    Private Role As Role


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        'ʵ������ɫ�û�����
        Role = New Role(conn, ts)


    End Sub

    '���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        If finishedFlag = "ͬ��" Then
            '��ȡ��Ŀ����A,B
            Dim tmpManagerA As String
            Dim strsql As String = "select manager_A from queryProjectInfo where ProjectCode='" & projectID & "'"

            Dim CommonQuery As New CommonQuery(conn, ts)
            Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strsql)

            '�쳣����  
            If dsProjectInfo.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
                Throw wfErr
            End If

            tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("manager_A")

            Dim TimingServer As New TimingServer(conn, ts, True, True)
            'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpManagerA, "28", "N")
            TimingServer.AddMsgContent(workFlowID, projectID, taskID, tmpManagerA, "��ͬ���ͨ��,����Ŀ�����ӡ��ͬ!", "N")

        End If
    End Function

End Class
