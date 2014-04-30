'判断是否收齐担保费
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsReturnGuaranteeFee
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private ProjectAccountDetail As ProjectAccountDetail

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

        ProjectAccountDetail = New ProjectAccountDetail(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        '获取担保费信息
        strSql = "{project_code='" & projectID & "' and item_type='31' and item_code='002' and income is not null}"
        Dim dsAccount, dsTemp As DataSet
        dsAccount = ProjectAccountDetail.GetProjectAccountDetailInfo(strSql)
        Dim i As Integer
        Dim sIncome As Single

        '计算担保费收取总额
        For i = 0 To dsAccount.Tables(0).Rows.Count - 1
            sIncome = sIncome + dsAccount.Tables(0).Rows(i).Item("income")
        Next

        '如果担保费未收
        ' 将IsReturnGuaranteeFee-ReturnGuaranteeFee置为假，IsReturnGuaranteeFee-SubmitCancelProArchives置为真。
        If sIncome = 0 Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='ReturnGuaranteeFee'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='SubmitCancelProArchives'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        Else

            '否则,将IsReturnGuaranteeFee-ReturnGuaranteeFee置为真，IsReturnGuaranteeFee-SubmitCancelProArchives置为假。

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='ReturnGuaranteeFee'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsReturnGuaranteeFee' and next_task='SubmitCancelProArchives'}"
            dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            dsTemp.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)

        End If

    End Function
End Class
