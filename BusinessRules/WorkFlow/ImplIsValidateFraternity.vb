
'判断是否需要互助会手续办妥确认
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsValidateFraternity
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private OppositeGuaranteeForm As OppositeGuaranteeForm
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

        OppositeGuaranteeForm = New OppositeGuaranteeForm(conn, ts)
        ProjectGuaranteeForm = New ProjectGuaranteeForm(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        Dim dsFraternity, dsGuarantee, dsTempTaskTrans, dsAttend As DataSet

        '①	在Opposite_Guarantee_Form表中获取互助会的数据字典定义；
        strSql = "{form_code='06'}"
        dsFraternity = OppositeGuaranteeForm.GetOppositeGuaranteeForm(strSql)
        Dim tmpFraternity As String = Trim(dsFraternity.Tables(0).Rows(0).Item("name"))

        '在Project_Guarantee_Form中获取
        strSql = "{project_code=" & "'" & projectID & "'" & " and ltrim(rtrim(guarantee_form))='" & tmpFraternity & "' and is_used=1}"
        dsGuarantee = ProjectGuaranteeForm.GetProjectGuaranteeForm(strSql)

        '②	如果获取的互助会对象为空
        If dsGuarantee.Tables(0).Rows.Count = 0 Then
            '    将IsValidateFraternity到ValidateFraternity的转移条件置为.F.;

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateFraternity' and next_task='ValidateFraternity'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr()
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            '    将ValidateFraternity的任务状态置为”F”;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateFraternity'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        Else
            '否则()
            '    将IsValidateFraternity到ValidateFraternity的转移条件置为.T.;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsValidateFraternity' and next_task='ValidateFraternity'}"
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
