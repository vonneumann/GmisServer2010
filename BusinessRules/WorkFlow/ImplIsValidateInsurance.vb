Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsValidateInsurance
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private ProjectGuaranteeForm As ProjectGuaranteeForm


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
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        ProjectGuaranteeForm = New ProjectGuaranteeForm(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim i As Integer
        Dim strSql As String
        '在Project_Guarantee_Form表中获取抵押物记录
        strSql = "{project_code=" & "'" & projectID & "'" & " and guarantee_form='抵押' and is_used=1}"
        Dim dsGuarantee, dsTempTaskTrans, dsAttend As DataSet
        dsGuarantee = ProjectGuaranteeForm.GetProjectGuaranteeForm(strSql)

        '如果抵押物记录为空
        If dsGuarantee.Tables(0).Rows.Count = 0 Then
            ' 将IsValidateInsurance到ValidateInsurance转移条件.F.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateInsurance' and next_task ='ValidateInsurance'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            ' 将ValidateInsurance的状态置为F
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id ='ValidateInsurance'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        Else
            '否则
            ' 将IsValidateInsurance到ValidateInsurance转移条件.T.
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateInsurance' and next_task='ValidateInsurance'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
        End If
    End Function
End Class
