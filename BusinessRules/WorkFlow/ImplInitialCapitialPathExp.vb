Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplInitialCapitialPathExp
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义工作记录对象引用
    Private WorkLog As WorkLog

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义定时任务对象引用
    Private WfProjectTimingTask As WfProjectTimingTask

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

        '实例化转移任务任务对象
        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim dsTempTaskTrans As DataSet
        '将评估资产任务（TID=CapitialEvaluated）到记录评审会结论的转移条件置为假.F.
        '将评估资产任务（TID=CapitialEvaluated）到登记反担保物的转移条件置为真.T.
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluatedExp' and next_task='RecordReviewConclusionExp'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '异常处理  
        If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
            Throw wfErr
        End If

        dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CapitialEvaluatedExp' and next_task='ApplyCapitialEvaluatedExp'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

        '异常处理  
        If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
            Throw wfErr
        End If

        dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

    End Function
End Class
