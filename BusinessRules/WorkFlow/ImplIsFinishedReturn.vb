'判断已经还款完毕
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsFinishedReturn
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private WfProjectTimingTask As WfProjectTimingTask

    Private ProjectAccountDetail As ProjectAccountDetail

    '定义支出金额、收入总额
    Private TrialFeePayout As Single = 0
    Private TotalTrialFeeIncome As Single = 0

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化转移任务对象
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        ProjectAccountDetail = New ProjectAccountDetail(conn, ts)

        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer

        '判断项目是否有损失
        strSql = "{project_code='" & projectID & "' and item_type='31' and item_code='004' and type='损失'}"
        Dim dsProjectAccountDetail As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim iCount As Integer = dsProjectAccountDetail.Tables(0).Rows.Count

        Dim dsTempTaskTrans As DataSet
        strSql = "{project_code=" & "'" & ProjectID & "'" & " and task_id='RefundRecord'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


        '如果还款未完毕并且没有损失
        If TrialFeePayout <> TotalTrialFeeIncome And iCount <> 0 Then

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RefundRecord" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else

            '否则

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RefundRecord" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        '将还款登记定时提示置为"E"
        strSql = "{project_code='" & projectID & "' and task_id='RefundRecord'}"
        Dim dsTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTimingTask.Tables(0).Rows.Count - 1
            dsTimingTask.Tables(0).Rows(i).Item("status") = "E"
        Next
        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTimingTask)

    End Function

    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '获取该项目关于评审费的记录
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code='" & ProjectID & "' item_type='31' and item_code='004' and type='还款'}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '计算评审费的收入总额
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout = TrialFeePayout + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout"))
            TotalTrialFeeIncome = TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income"))
        Next
    End Function
End Class
