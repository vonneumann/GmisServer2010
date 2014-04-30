
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplAddValuator
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义消息对象引用
    Private WfProjectMessages As WfProjectMessages

    '定义通用查询对象引用
    Private CommonQuery As CommonQuery

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化参与人对象
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        '实例化消息对象
        WfProjectMessages = New WfProjectMessages(conn, ts)

        '实例化通用查询对象
        CommonQuery = New CommonQuery(conn, ts)



    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '获取该项目的资产评估师
        Dim strSql As String
        Dim i, j As Integer
        Dim dsTempValuator, dsTempAttend, dsTempDeleteAttend, dsTempRoleTemplate, dsTempTaskMessages, dsTempGuarantyName As DataSet
        Dim iAttendCount, iValuatorCount As Integer
        Dim tmpValuate_person, tmpValuateGuarantyType, tmpValuateGuarantyName, tmpWorkflowID As String
        Dim newRow As DataRow
        Dim itemType As New Item(conn, ts)
        strSql = "SELECT DISTINCT evaluate_person" & _
                 " FROM opposite_guarantee" & _
                 " WHERE evaluate_person is not null" & _
                 " and project_code =" & "'" & projectID & "'"
        dsTempValuator = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsTempValuator.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempValuator.Tables(0))
            Throw wfErr
        End If

        iValuatorCount = dsTempValuator.Tables(0).Rows.Count

        '先获取当前任务的模版ID(作为复制模版的实例时的限制)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
        If dsTempDeleteAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempDeleteAttend.Tables(0))
            Throw wfErr
        End If

        tmpWorkflowID = dsTempDeleteAttend.Tables(0).Rows(0).Item("workflow_id")

        '先删除任务角色表中原有资产评估师的任务（上一次分配的评估师）
        'strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'" & "}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'" & " and workflow_id='" & tmpWorkflowID & "'}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        For i = 0 To dsTempDeleteAttend.Tables(0).Rows.Count - 1
            dsTempDeleteAttend.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        ''AssignValuator_Update
        ''假如分配完评估师提交过去，然后启动更新反担保措施任务，重新分配评估师，并分配给另外的评估师，以前所分配的评估师任务记录要删除掉
        'If taskID = "AssignValuator_Update" Then
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'" & " and task_id='" & tmpWorkflowID & "'}"
        '    dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    For i = 0 To dsTempDeleteAttend.Tables(0).Rows.Count - 1
        '        dsTempDeleteAttend.Tables(0).Rows(i).Delete()
        '    Next
        '    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)
        'End If



        '再复制任务角色模版的资产评估师的任务
        strSql = "{workflow_id=" & "'" & tmpWorkflowID & "'" & " and role_id='34'" & "}"
        Dim WfTaskRoleTemplate As New WfTaskRoleTemplate(conn, ts)
        dsTempRoleTemplate = WfTaskRoleTemplate.GetWfTaskRoleTemplateInfo(strSql)

        For i = 0 To dsTempRoleTemplate.Tables(0).Rows.Count - 1
            newRow = dsTempDeleteAttend.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = tmpWorkflowID
                .Item("project_code") = projectID    '将所有添加任务的工作流ID置为项目编码
                .Item("task_id") = dsTempRoleTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTempRoleTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTempDeleteAttend.Tables(0).Rows.Add(newRow)
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        ' strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34' and workflow_id='" & tmpWorkflowID & "'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        iAttendCount = dsTempAttend.Tables(0).Rows.Count

        '④	对于每位资产评估师
        ' 在项目的任务参与人表复制role-id=34的任务（n-1） ,将参与人改为资产评估师；
        strSql = "evaluate_person desc"
        dsTempValuator.Tables(0).Select("", strSql)
        Dim tmpPreValuator As String
        If iValuatorCount <> 0 Then


            '如果只有一个评估师，则把每个角色任务的参与人置为该评估师

            For i = 0 To iAttendCount - 1

                dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsTempValuator.Tables(0).Rows(0).Item("evaluate_person")

            Next

            '如果有多个评估师，则为其他的评估师分配每个角色任务

            If iValuatorCount >= 2 Then

                For i = 1 To iValuatorCount - 1

                    For j = 0 To dsTempRoleTemplate.Tables(0).Rows.Count - 1
                        newRow = dsTempAttend.Tables(0).NewRow
                        With dsTempRoleTemplate.Tables(0).Rows(j)
                            newRow.Item("project_code") = projectID
                            newRow.Item("workflow_id") = .Item("workflow_id")
                            newRow.Item("task_id") = .Item("task_id")
                            newRow.Item("role_id") = .Item("role_id")
                            newRow.Item("attend_person") = dsTempValuator.Tables(0).Rows(i).Item("evaluate_person")
                        End With
                        dsTempAttend.Tables(0).Rows.Add(newRow)

                    Next

                Next
             
            End If

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        Else

            '抛出未分配资产评估师错误
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowAddValuatorErr()
            Throw wfErr

        End If



    End Function

End Class
