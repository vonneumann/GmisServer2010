
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendMsgToManagerA
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义定时服务对象引用
    Private TimingServer As TimingServer

    '定义通用查询对象引用
    Private CommonQuery As CommonQuery


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化定时服务对象引用
        TimingServer = New TimingServer(conn, ts, True, True)

        '实例化通用查询对象
        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    '启动收取评审费时通知项目经理收取评审费
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '获取项目经理A,B
        Dim tmpManagerA As String
        Dim strsql As String = "select manager_A from queryProjectInfo where ProjectCode='" & projectID & "'"
        Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strsql)

        '异常处理  
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
            strAffirmEvaluateDate = "确认评估日期为：" & dsOpp.Tables(0).Rows(0).Item("affirm_evaluate_date")
        End If
        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpManagerA, "28", "N")
        TimingServer.AddMsgContent(workFlowID, projectID, taskID, tmpManagerA, strAffirmEvaluateDate, "N")

    End Function



End Class
