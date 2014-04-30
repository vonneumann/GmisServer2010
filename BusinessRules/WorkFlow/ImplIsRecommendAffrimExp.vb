Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'确认合作区
Public Class ImplIsRecommendAffrimExp
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义转移任务对象引用
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private CooperateOpinion As CooperateOpinion


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

        CooperateOpinion = New CooperateOpinion(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer
        Dim dsCooperate, dsTempTaskTrans, dsAttend As DataSet
        '①	在Cooperate-Organization-Opinion表获取项目的合作区对象Cooperate-Organization；
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        dsCooperate = CooperateOpinion.GetCooperateOpinionInfo("null", strSql)

        '②	如果获取的合作区对象为空
        If dsCooperate.Tables(1).Rows.Count = 0 Then
            '    将IsRecommendAffrim到RecommendAffrim的转移条件置为.F.;

            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsRecommendAffrimExp' and next_task='RecommendAffrimExp'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".F."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            '    将RecommendAffrim的任务状态置为”F”;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecommendAffrimExp'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("task_status") = "F"
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        Else
            '否则()
            '    将IsRecommendAffrim到RecommendAffrim的转移条件置为.T.;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='IsRecommendAffrimExp' and next_task='RecommendAffrimExp'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)

            '异常处理  
            If dsTempTaskTrans.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTempTaskTrans.Tables(0))
                Throw wfErr
            End If

            dsTempTaskTrans.Tables(0).Rows(0).Item("transfer_condition") = ".T."
            WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)

            Dim tmpCooperate As String = dsCooperate.Tables(1).Rows(0).Item("cooperate_organization")
            '    在Cooperate-Organization表获取Cooperate-Organization对象的专管员Manager;
            strSql = "{cooperate_organization=" & "'" & tmpCooperate & "'" & "}"
            dsCooperate = CooperateOpinion.GetCooperateOpinionInfo(strSql, "null")
            'qxd modify
            'Dim tmpManager As String = dsCooperate.Tables(0).Rows(0).Item("manager")
            Dim tmpManager As String = IIf(dsCooperate.Tables(0).Rows(0).Item("manager") Is System.DBNull.Value, "", dsCooperate.Tables(0).Rows(0).Item("manager"))

            '    将RecommendAffrim任务的参与人置为Manager;
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecommendAffrimExp'}"
            dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

            For i = 0 To dsAttend.Tables(0).Rows.Count - 1
                dsAttend.Tables(0).Rows(i).Item("attend_person") = tmpManager
            Next

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        End If
    End Function
End Class
