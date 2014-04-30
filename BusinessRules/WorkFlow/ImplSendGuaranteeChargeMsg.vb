Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendGuaranteeChargeMsg
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
        Dim strsql As String = "{ProjectCode=" & "'" & projectID & "'" & "}"
        Dim dsProjectInfo As DataSet = CommonQuery.GetProjectInfoEx(strsql)

        '异常处理  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("24")

        TimingServer.AddMsg(workFlowID, projectID, taskID, tmpManagerA, "22", "N")

    End Function
End Class
