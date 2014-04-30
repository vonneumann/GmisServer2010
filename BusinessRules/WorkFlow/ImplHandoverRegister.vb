Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplHandoverRegister
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private WfProjectTaskAttendee As WfProjectTaskAttendee

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

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim tmpAttend As String

        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='Review'}"
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        If dsTemp.Tables(0).Rows.Count = 0 Then

            '����Ϣ����Ŀ������Ա
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='Application'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
            TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpAttend, "13", "N")

        Else
            tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")

            '��������������˲��գ�����Ŀ�ƽ���Ϣ������Ŀ������Ա
            If tmpAttend <> "" Then
                TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpAttend, "13", "N")
            Else
                '���򷢸���Ŀ������Ա
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='Application'}"
                dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '�쳣����  
                If dsTemp.Tables(0).Rows.Count = 0 Then
                    Dim wfErr As New WorkFlowErr()
                    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                    Throw wfErr
                End If

                tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
                TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpAttend, "13", "N")
            End If
        End If

    End Function

End Class
