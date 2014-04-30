Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports BusinessRules.OAWorkflowXYDB

'删除修改记录评审会流程
Public Class ImplEndModifyConclusion
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WfProjectTask As WfProjectTask
    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTimingTask As WfProjectTimingTask
    Private WfProjectTrack As WfProjectTrack


    Private OAWorkflowXYDB As OAWorkflowXYDB.WorkflowServiceForXYDB
    Private webserviceCgmisForOA As WebserviceCgmisForOA.ServiceOA

    Private commonquery As CommonQuery

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        WfProjectTask = New WfProjectTask(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        WfProjectTrack = New WfProjectTrack(conn, ts)


        OAWorkflowXYDB = New OAWorkflowXYDB.WorkflowServiceForXYDB()
        webserviceCgmisForOA = New WebserviceCgmisForOA.ServiceOA()

        commonquery = New CommonQuery(conn, ts)

    End Sub


    '创建暂缓子流程
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        strSql = "{project_code='" & projectID & "' and workflow_id='15'}"
        Dim dsTemp As DataSet



        '删除参与人表
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        '删除转移表
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        '删除定时任务表
        dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        '删除跟踪表
        dsTemp = WfProjectTrack.GetWfProjectTrackInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTrack.UpdateWfProjectTrack(dsTemp)


        '删除任务表
        dsTemp = WfProjectTask.GetWfProjectTaskInfo(strSql)

        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Delete()
        Next

        WfProjectTask.UpdateWfProjectTask(dsTemp)


        '2010-05-13 yjf add 设置评审会通过后长期未签约预警消息
        'If workFlowID <> "08" And workFlowID <> "10" Then

        strSql = "{project_code is null}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        Dim newRow As DataRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = "02"
            .Item("project_code") = projectID
            .Item("task_id") = "RecordSignature"
            .Item("role_id") = "24"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 30, Now)
            .Item("status") = "P"
            .Item("time_limit") = 30
            .Item("distance") = 0
            .Item("message_id") = 32
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = "02"
            .Item("project_code") = projectID
            .Item("task_id") = "RecordSignature"
            .Item("role_id") = "29"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 30, Now)
            .Item("status") = "P"
            .Item("time_limit") = 30
            .Item("distance") = 0
            .Item("message_id") = 32
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = "02"
            .Item("project_code") = projectID
            .Item("task_id") = "RecordSignature"
            .Item("role_id") = "21"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 30, Now)
            .Item("status") = "P"
            .Item("time_limit") = 30
            .Item("distance") = 0
            .Item("message_id") = 32
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        newRow = dsTempTimingTask.Tables(0).NewRow
        With newRow
            .Item("workflow_id") = "02"
            .Item("project_code") = projectID
            .Item("task_id") = "RecordSignature"
            .Item("role_id") = "02"
            .Item("type") = "M"
            .Item("start_time") = DateAdd(DateInterval.Day, 30, Now)
            .Item("status") = "P"
            .Item("time_limit") = 30
            .Item("distance") = 0
            .Item("message_id") = 32
        End With
        dsTempTimingTask.Tables(0).Rows.Add(newRow)

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        'End If     
    End Function
End Class
