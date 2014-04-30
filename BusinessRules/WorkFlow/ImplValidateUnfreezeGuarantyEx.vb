
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplValidateUnfreezeGuarantyEx
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
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

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '项目采用反担保措施了,则启动解除反担保措施任务,并发送消息;否则启动项目终止

        '定义放款额、还款额
        Dim TrialFeePayout As Double = 0.0
        Dim TotalTrialFeeIncome As Double = 0.0
        Dim DepositFee As Double = 0.0

        TrialFeePayout = getTotalLoan(projectID)
        TotalTrialFeeIncome = getTotalTrialFeeIncome(projectID)
        DepositFee = getDepositFee(projectID)

        Dim strSql As String

        Dim i As Integer
        Dim dsTempTaskTrans, dsAttend As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectFinishedRrport'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '当还款完成并且保证金已收取
        'If TotalTrialFeeIncome >= TrialFeePayout And DepositFee > 0.0 Then
        If DepositFee > 0.0 Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                'If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "VaildateUnfreezeGuaranty" Then
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateReturnDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ValidateReturnDepositFee" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                ElseIf dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ProjectFinished" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next

        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    '获得确认收取保证金的金额
    Private Function getDepositFee(ByVal ProjectID As String) As Double

        '获取该项目关于保证金的记录(34,009)
        Dim DepositFee As Double = 0.0
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and item_type='34'" & " and item_code='009'" & "}"
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer

        '计算保证金的收入总额
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            DepositFee = CDbl(DepositFee + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("payout")), 0, dsTemp.Tables(0).Rows(i).Item("payout")))
        Next

        Return DepositFee

    End Function

    '获取该项目关于放款额的金额
    Private Function getTotalLoan(ByVal ProjectID As String) As Double

        Dim TotalLoan As Double = 0.0

        '获取该项目关于放款额的记录
        Dim ProjectLoanNotice As New LoanNotice(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'}"
        Dim dsTemp As DataSet = ProjectLoanNotice.GetLoanNoticeInfo(strSql)
        Dim i As Integer

        '计算放款总额
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalLoan = CDbl(TotalLoan + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("sum")), 0, dsTemp.Tables(0).Rows(i).Item("sum")))
            Next
        End If

        Return TotalLoan
    End Function

    '获得并计算还款额总额
    Private Function getTotalTrialFeeIncome(ByVal ProjectID As String) As Double
        Dim strSql As String
        Dim i As Integer

        '获得并计算还款额总额
        Dim TotalTrialFee As Double = 0.0
        strSql = "{project_code=" & "'" & ProjectID & "'" & " and item_type='34' and item_code='001'}"
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim TotalTrialFeeIncome As Double '还款总额

        '还款总额
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalTrialFee = CDbl(TotalTrialFee + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
            Next
        End If

        Return TotalTrialFee
    End Function
End Class
