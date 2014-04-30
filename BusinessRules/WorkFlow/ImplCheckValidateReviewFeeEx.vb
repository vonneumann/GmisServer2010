Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplCheckValidateReviewFeeEx
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义评审费支出金额、收入总额
    Private TrialFeePayout As Integer = 0
    Private TotalTrialFeeIncome As Integer = 0

    '定义项目意见对象引用
    Private attendee As ProjectTaskAttendee


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化项目意见对象
        attendee = New ProjectTaskAttendee(conn, ts)

    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '如果该项目的反担保措施包括互助协会承担30%则不需启动确认收费标准任务
        Dim projectGuaranteeForm As New ProjectGuaranteeForm(conn, ts)
        Dim strSql As String
        strSql = "{project_code='" & projectID & "' and guarantee_form like '%互助协会%' and isnull(is_used,0)=1}"
        Dim tmpDataSet As DataSet
        tmpDataSet = projectGuaranteeForm.GetProjectGuaranteeForm(strSql)
        Dim hasHuZhuHui As Boolean = (tmpDataSet.Tables(0).Rows.Count > 0)
        tmpDataSet.Dispose()

      

        Dim hasDoneConfirmReviewFee As Boolean
        Dim payoutIsBigger As Boolean

        '互助会项目不需要收取评审费，否则需要计算是否要受评审费。
        If hasHuZhuHui = False Then
            Dim dsAttendee As DataSet = attendee.GetProjectTaskAttendeeInfo("{project_code='" & projectID & "' AND task_id='BalanceReviewFee'}")
            If dsAttendee.Tables(0).Rows.Count > 0 Then
                hasDoneConfirmReviewFee = CBool(dsAttendee.Tables(0).Rows(0)("task_status") & "" = "F")
            End If
            dsAttendee.Dispose()

            '计算评审费的收入总额
            GetTotalTrialFeeIncome(projectID)

            payoutIsBigger = (TotalTrialFeeIncome < TrialFeePayout)
        End If

        Dim i As Integer
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckValidateReviewFee'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
            If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "SendCashlossReviewMsg" Then
                dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = IIf(hasHuZhuHui, ".F.", IIf(hasDoneConfirmReviewFee AndAlso payoutIsBigger, ".T.", ".F."))
            Else
                dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = IIf(hasHuZhuHui, ".T.", IIf(hasDoneConfirmReviewFee AndAlso payoutIsBigger, ".F.", ".T."))
            End If
        Next

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function



    Private Function GetTotalTrialFeeIncome(ByVal ProjectID As String)

        '获取该项目关于评审费的记录
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='31'" & " and item_code='001'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '计算评审费的收入总额
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            TrialFeePayout += CDbl(IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout")))
            TotalTrialFeeIncome += CDbl(IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
        Next
    End Function
End Class
