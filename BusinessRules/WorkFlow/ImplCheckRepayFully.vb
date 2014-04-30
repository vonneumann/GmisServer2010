Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplCheckRepayFully
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获得该项目的放款总额
        Dim TotalLoanIncome As Double
        TotalLoanIncome = getTotalLoan(projectID) * 10000.0 '单位万元转化为：元

        '判断还款额是否大于等于放款额；是：则流转到登记还款证明书；否则：流转到登记还款
        Dim strSql As String
        Dim i As Integer

        '获得并计算还款额总额
        strSql = "{project_code=" & "'" & projectID & "'" & " and item_type='34' and item_code='001'}"
        Dim ProjectAccountDetail As New ProjectAccountDetail(conn, ts)
        Dim dsTemp As DataSet = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim TotalTrialFeeIncome As Double '还款总额

        '还款总额
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalTrialFeeIncome = CDbl(TotalTrialFeeIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("income")), 0, dsTemp.Tables(0).Rows(i).Item("income")))
            Next
        End If

        '获得并设置转移条件
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CHKRepayFully'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        If TotalTrialFeeIncome >= TotalLoanIncome Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RecordRefundCertificate" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "RecordRefundCertificate" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function

    '
    Private Function getTotalLoan(ByVal ProjectID As String) As Double

        Dim TotalLoanIncome As Double = 0.0

        '获取该项目关于放款额的记录
        Dim ProjectLoanNotice As New LoanNotice(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'}"
        Dim dsTemp As DataSet = ProjectLoanNotice.GetLoanNoticeInfo(strSql)
        Dim i As Integer

        '计算放款总额
        If dsTemp.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                TotalLoanIncome = CDbl(TotalLoanIncome + IIf(IsDBNull(dsTemp.Tables(0).Rows(i).Item("sum")), 0, dsTemp.Tables(0).Rows(i).Item("sum")))
            Next
        End If

        Return TotalLoanIncome
    End Function
End Class
