Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'再担保中记录评审会评审结论：判断结论

Public Class ImplMeetingRecord
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WfProjectTaskTransfer As WfProjectTaskTransfer

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
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

    End Function
End Class
