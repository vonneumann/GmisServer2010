Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


'把有固定员工的角色的员工ID添加到参与人中
Public Class ImplSetManager
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义角色用户对象引用
    Private Role As Role


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

        '实例化角色用户对象
        Role = New Role(conn, ts)


    End Sub

    '把有固定员工的角色的员工ID添加到参与人中
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获取该项目的工作任务参与人信息
        Dim strAttendee As String
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id='" & "ReviewMeetingPlan" & "'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
        If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            Throw wfErr
        End If

        If Not dsTempTaskAttendee Is Nothing Then
            strAttendee = dsTempTaskAttendee.Tables(0).Rows(0).Item("attend_person")
        End If

        '''''''''''''''
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='" & "RecordReviewConclusion" & "'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        Dim i As Integer
        Dim dsTempRole As DataSet
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1

            '把有固定员工的角色的员工ID添加到参与人中
            With dsTempTaskAttendee.Tables(0).Rows(i)
                .Item("attend_person") = strAttendee

                Select Case Trim(.Item("role_id"))
                    Case "01" '中心主任
                        dsTempRole = Role.GetStaffRole("01")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "12" '资信部长
                        dsTempRole = Role.GetStaffRole("12")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "21" '担保部长
                        dsTempRole = Role.GetStaffRole("21")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "31" '风险部长
                        dsTempRole = Role.GetStaffRole("31")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "40" '综合管理部长
                        dsTempRole = Role.GetStaffRole("40")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "41" '财务人员
                        dsTempRole = Role.GetStaffRole("41")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "14" '初审统计人员
                        dsTempRole = Role.GetStaffRole("14")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "27" '项目登记员
                        dsTempRole = Role.GetStaffRole("27")

                        '异常处理  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")

                End Select


            End With
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

    End Function

End Class
