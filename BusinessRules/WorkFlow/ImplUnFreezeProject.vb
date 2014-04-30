'项目解冻
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


Public Class ImplUnFreezeProject
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WfProjectTaskAttendee As WfProjectTaskAttendee
    Private WfProjectTimingTask As WfProjectTimingTask
    Private workflow As WorkFlow

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)
        workflow = New WorkFlow(conn, ts)

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获取项目编码前5位与本项目的项目编码前5位相同， RefundRecord或登记还款证明书任务状态为P的项目数PN和项目编码数PN；
        Dim sProjectCode As String = Mid(projectID, 1, 5)
        Dim strSql As String = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RefundRecord','RecordRefundCertificate') and task_status='P'}"
        Dim dsTemp, dsProjectNum As DataSet
        dsProjectNum = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim iPN, i, j, k As Integer
        iPN = dsProjectNum.Tables(0).Rows.Count

        '如果PN=0，将项目编码前5位与本项目的项目编码前5位相同、登记保后检查记录表、记录保后跟踪活动、项目评价、审核保后跟踪检查记录表,复核保后检查记录表任务状态为‘P‘的任务置为F，
        '停止以上任务定时服务；
        If iPN = 0 Then
            strSql = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess','CheckProjectProcess','CheckProjectTraceInfo') and task_status='P'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            strSql = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess','CheckProjectProcess','CheckProjectTraceInfo') and status='P'}"
            dsTemp = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("status") = "E"
            Next
            WfProjectTimingTask.UpdateWfProjectTimingTask(dsTemp)

        Else
            ''否则只保留第一个项目的保后跟踪活动（将其他项目编码前5位与本项目的项目编码前5位相同、登记保后检查记录表、记录保后跟踪活动、项目评价、审核保后跟踪检查记录表任务状态为‘P‘的任务置为‘’）。
            'strSql = "{substring(project_code,1,5)='" & sProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess','CheckProjectProcess','CheckProjectTraceInfo') and task_status='P'}"
            'dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            'Dim tmpProjectCode As String = dsTemp.Tables(0).Rows(0).Item("project_code") '获取第一个项目的项目编码

            'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            '    If dsTemp.Tables(0).Rows(i).Item("project_code") <> tmpProjectCode Then
            '        dsTemp.Tables(0).Rows(i).Item("task_status") = ""
            '    End If
            'Next

            'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            '2005-3-15 yjf edit 修改项目解冻功能：确保企业至少有一个项目有保后跟踪活动(找到的第一个项目)
            Dim tmpProjectCodeNumOne, tmpProjectCode, tmpTaskID As String
            tmpProjectCodeNumOne = dsProjectNum.Tables(0).Rows(0).Item("project_code") '获取第一个项目的项目编码

            strSql = "{project_code='" & tmpProjectCodeNumOne & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess')}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                tmpTaskID = dsTemp.Tables(0).Rows(i).Item("task_id")
                workflow.StartupTask(workFlowID, tmpProjectCodeNumOne, tmpTaskID, "", "")
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

            '确保其他项目的保后跟踪都关闭
            For j = 0 To dsProjectNum.Tables(0).Rows.Count - 1
                tmpProjectCode = dsProjectNum.Tables(0).Rows(j).Item("project_code")
                strSql = "{project_code='" & tmpProjectCode & "' and task_id in ('RecordProjectTraceInfo','RecordProjectProcess','AppraiseProjectProcess')}"
                dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
                For k = 0 To dsTemp.Tables(0).Rows.Count - 1
                    If dsTemp.Tables(0).Rows(k).Item("project_code") <> tmpProjectCodeNumOne Then
                        dsTemp.Tables(0).Rows(k).Item("task_status") = ""
                    End If
                Next
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

        End If
    End Function
End Class
