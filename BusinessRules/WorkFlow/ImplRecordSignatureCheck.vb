
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplRecordSignatureCheck
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义消息对象引用
    Private WfProjectMessages As WfProjectMessages

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

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        '实例化消息对象
        WfProjectMessages = New WfProjectMessages(conn, ts)

        '实例化通用查询对象
        CommonQuery = New CommonQuery(conn, ts)



    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        If userID <> "周新红" Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowMustRecordSignatureSubmitor()
            Throw wfErr
        End If
    End Function

End Class
