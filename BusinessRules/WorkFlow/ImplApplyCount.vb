Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplApplyCount
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
        '①	在担保意向书（itent_letter）获取申请次数；
        Dim applyCount As Integer = GetApplyCount(projectID)

        '②	如果申请次数小于3，将转移条件FROMID=65、TOID=24 的转移条件置为“.T.”，将FROMID=65、TOID=62 的转移条件置为“.F.”;
        '否则，将FROMID=65、TOID=24 的转移条件置为“.F.”，将FROMID=65、TOID=62 的转移条件置
        '为“.T.”;
        Dim strSql As String
        Dim i As Integer
        Dim dsTempTaskTrans As DataSet
        Dim WfProjectTaskTransfer As New WfProjectTaskTransfer(conn, ts)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckApplyTimes'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
        If applyCount < _ApplyNumLimit Then
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ApplyLetterIntent" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else
            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "ApplyLetterIntent" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)


        '更新项目信息


    End Function

    '获取申请次数
    Private Function GetApplyCount(ByVal ProjectID As String) As Integer
        Dim IntentLetter As New IntentLetter(conn, ts)
        'qxd modify 2004-9-24
        'Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & "}"
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & " and bank_reply='不同意'}"
        Dim dsTemp As DataSet = IntentLetter.GetIntentLetterInfo(strSql)

        '记录的条数即申请的次数
        Return dsTemp.Tables(0).Rows.Count

    End Function
End Class
