Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'结束项目评审任务
Public Class ImplEndProjectAppraise
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
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


        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

    End Sub

    '关闭任务:项目评审,记录保前调研活动,登记反担保措施,提交评审意见,提交调研结论
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        'strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ProjectAppraiseReport'" & "}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('ProjectAppraiseReport','PreguaranteeActivity','ApplyCapitialEvaluated','ProjectAttitude','SubmissionProbeResult')" & "}"
        Dim dsTempAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempAttend.Tables(0).Rows.Count - 1
            dsTempAttend.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)
    End Function
End Class
