Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendReviewFeeChargeMsg
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    ''定义项目对象引用
    'Private project As project

    ''定义定时服务对象引用
    'Private TimingServer As TimingServer


    ''定义通用查询对象引用
    'Private CommonQuery As CommonQuery


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        ''实例化项目对象
        'project = New Project(conn, ts)

        ''实例化定时服务对象引用
        'TimingServer = New TimingServer(conn, ts, True, True)

        ''实例化通用查询对象
        'CommonQuery = New CommonQuery(conn, ts)

    End Sub

    '启动收取评审费时通知项目经理收取评审费
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        ''获取初审人员
        'Dim strsql As String = "{ProjectCode=" & "'" & projectID & "'" & "}"
        'Dim dsProjectInfo As DataSet = CommonQuery.GetProjectInfoEx(strsql)

        ''异常处理  
        'If dsProjectInfo.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr()
        '    wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
        '    Throw wfErr
        'End If

        'Dim tmpAttend As String = dsProjectInfo.Tables(0).Rows(0).Item("13")
        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpAttend, "20", "N")


        '2010-05-13 yjf add 设置项目受理后长期未预交评审费预警消息
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
