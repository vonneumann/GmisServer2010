Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'在定时任务表将与参数匹配的提示任务状态置为“E”；
Public Class ImplDeleteTimingAlert
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

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

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '①	在定时任务表将与参数匹配的提示任务状态置为“E”；
        Dim strSql As String
        Dim dsTempTimingTask As DataSet
        Dim i As Integer
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsTempTimingTask = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)

        '   所有地提示任务状态置为“E”
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

    End Function


End Class
