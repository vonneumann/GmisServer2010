Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplVaildateUnfreezeGuaranty
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WorkFlow As Workflow

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

        Workflow = New WorkFlow(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '向档案管理人员和风险管理部长发办理解除反担保物消息
        Dim tmpFileManager, tmpMinister As String
        tmpFileManager = WorkFlow.getTaskActor("42")
        tmpMinister = WorkFlow.getTaskActor("31")

        TimingServer.AddMsg(workFlowID, projectID, taskID, tmpFileManager, "23", "N")
        TimingServer.AddMsg(workFlowID, projectID, taskID, tmpMinister, "23", "N")
    End Function
End Class
