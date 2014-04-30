Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplHandoverRegister
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private TimingServer As TimingServer

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
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

            '发消息给项目受理人员
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='Application'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
            TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpAttend, "13", "N")

        Else
            tmpAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")

            '当初审任务参与人不空，将项目移交消息发往项目初审人员
            If tmpAttend <> "" Then
                TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpAttend, "13", "N")
            Else
                '否则发给项目受理人员
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='Application'}"
                dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                '异常处理  
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
