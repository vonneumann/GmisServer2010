'判断是否需要机构审核
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplIsAudit
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer

    Private project As Project

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

        project = New Project(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer

        '判断项目的申请金额是否大于50万
        strSql = "{project_code='" & projectID & "'}"
        Dim dsProject As DataSet = project.GetProjectInfo(strSql)
        Dim applySum As Single = IIf(IsDBNull(dsProject.Tables(0).Rows(0).Item("apply_sum")), 0, dsProject.Tables(0).Rows(0).Item("apply_sum"))

        Dim dsTempTaskTrans As DataSet
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsAudit'}"
        dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)


        '如果大于50万
        If applySum > 50 Then

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "Audit" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                End If
            Next
        Else

            '否则

            For i = 0 To dsTempTaskTrans.Tables(0).Rows.Count - 1
                If dsTempTaskTrans.Tables(0).Rows(i).Item("next_task") = "Audit" Then
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".F."
                Else
                    dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                End If
            Next
        End If

        WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)


    End Function

End Class
