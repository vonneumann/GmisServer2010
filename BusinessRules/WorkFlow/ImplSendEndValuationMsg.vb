Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendEndValuationMsg
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

    '定义参与人对象引用
    Private ProjectTaskAttendee As ProjectTaskAttendee

    '定义消息对象引用
    Private WfProjectMessages As WfProjectMessages


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化定时任务对象
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)


        '实例化参与人对象
        ProjectTaskAttendee = New ProjectTaskAttendee(conn, ts)


        '实例化消息对象
        WfProjectMessages = New WfProjectMessages(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获取项目经理
        Dim strSql As String
        Dim i As Integer
        Dim dsTempAttend, dsTempTaskMessages As DataSet
        Dim newRow As DataRow
        Dim tmpManager As String
        Dim drManager() As DataRow
        dsTempAttend = ProjectTaskAttendee.GetProjectAttendeeInfo(projectID)
        drManager = dsTempAttend.Tables(0).Select("role_id in ('24','25')")
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")

        '获取项目的企业名称
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim Project As New Project(conn, ts)
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)

        '异常处理  
        If dsProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProject.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")

        '异常处理  
        If dsCorporation.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsCorporation.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        For i = 0 To drManager.Length - 1
            tmpManager = drManager(i).Item("attend_person")
            newRow = dsTempTaskMessages.Tables(0).NewRow
            With newRow
                .Item("project_code") = projectID
                .Item("message_content") = userID & " " & tmpCorporationName & "项目" & "资产评估完毕"
                .Item("accepter") = tmpManager
                .Item("send_time") = Now
                .Item("is_affirmed") = "N"
            End With
            dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        Next
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)

    End Function
End Class
