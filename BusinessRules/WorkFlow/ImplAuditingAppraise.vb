Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplAuditingAppraise
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    'Private WorkFlow As WorkFlow
    'Private TimingServer As TimingServer
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        'WorkFlow = New WorkFlow(conn, ts)

        'TimingServer = New TimingServer(conn, ts, True, True)

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获得并设置转移条件
        Dim strSql As String
        Dim i As Integer
        Dim ds, dsTempTaskTrans As DataSet

        If finishedFlag = "否定" Then
            closeTaskStatus(projectID, taskID, userID)
        ElseIf finishedFlag = "上会" Then
            closeTaskStatus(projectID, taskID, userID)
        Else
            strSql = "{project_code ='" & projectID & "' and task_id='" & taskID & "' and task_status='F'}"
            ds = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            If ds.Tables(0).Rows.Count >= 3 Then
                strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='AuditingAppraise_AUTO'}"
                dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
                If dsTempTaskTrans.Tables(0).Rows.Count > 0 Then
                    For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                        If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "AuditingAppraise_All" Then
                            dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                        Else
                            dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                        End If
                    Next
                    WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
                End If
            End If
        End If
    End Function

    '关闭其他人的相同名称的任务
    Private Sub closeTaskStatus(ByVal projectCode As String, ByVal taskID As String, ByVal userID As String)
        Dim strSql As String
        Dim ds As DataSet
        Dim i As Integer

        strSql = "{project_code ='" & projectCode & "' and task_id='" & taskID & "' and attend_person<>'" & userID & "'}"
        ds = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If ds.Tables(0).Rows.Count > 0 Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                ds.Tables(0).Rows(i).Item("task_status") = "F"
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(ds)
        End If
    End Sub

End Class
