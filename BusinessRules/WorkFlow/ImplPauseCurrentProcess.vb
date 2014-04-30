'暂停当前正在处理的任务(除暂缓流程)
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplPauseCurrentProcess
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

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

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

       
    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer

        '如果是16（修改反担保措施流程），则评审阶段修改反担保措施不要将正在进行的任务挂起
        If workFlowID = "16" Then
            Dim strTemp As String
            Dim strPhase As String

            Dim commQuery As CommonQuery = New CommonQuery(conn, ts)
            strTemp = "select phase from project where project_code='" & projectID & "'"
            Dim ds As DataSet = commQuery.GetCommonQueryInfo(strTemp)
            If Not ds Is Nothing Then
                strPhase = IIf(ds.Tables(0).Rows(0).Item("phase") Is System.DBNull.Value, "", ds.Tables(0).Rows(0).Item("phase"))
                If strPhase = "评审" Then
                    Exit Function
                End If
            End If
        End If

        '将除工作流ID为13、17,22以外的所有的任务状态为P 的置为
        strSql = "{project_code='" & projectID & "' and workflow_id not in ('13','17','22') and task_status='P'}"
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Item("task_status") = "C"
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

    End Function
End Class
