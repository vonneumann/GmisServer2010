Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsRecordBankReply
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
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

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)



    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        Dim dsTemp As DataSet
        '判断RecordBankReply登记银行恢复任务是否已做过
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordBankReply'}"
        dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim tmpStatus As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("task_status")), "", dsTemp.Tables(0).Rows(0).Item("task_status"))

        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateReviewConclusion'}"
        dsTemp = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


        '如果做过则
        If tmpStatus = "F" And isAllowLaon(projectID) Then
            '          将ValidateReviewConclusion到NewReviewConclusion的转移条件置为.T.
            '          将ValidateReviewConclusion到其他任务的转移条件置为.F.
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                If dsTemp.Tables(0).Rows(i).Item("next_task") = "NewReviewConclusion" Then
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        Else
            '否则
            '          将ValidateReviewConclusion到NewReviewConclusion的转移条件置为.F.
            '          将ValidateReviewConclusion到其他任务的转移条件置为.T.
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                If dsTemp.Tables(0).Rows(i).Item("next_task") = "NewReviewConclusion" Then
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTemp.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTemp)
    End Function

    'qxd add 2004-9-24
    '获得Intent_letter的银行回复意见是否有一条记录为“同意”（注：一个项目最多只有一条记录为同意）。同意则允许放款
    Private Function isAllowLaon(ByVal projectID As String) As Boolean
        Dim strSql As String
        Dim intentLetter As IntentLetter = New IntentLetter(conn, ts)
        Dim ds As DataSet

        strSql = "{project_code='" & projectID & "' and bank_reply='同意'}"
        ds = intentLetter.GetIntentLetterInfo(strSql)
        If Not ds Is Nothing Then
            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
            Return False
        End If
    End Function
End Class
