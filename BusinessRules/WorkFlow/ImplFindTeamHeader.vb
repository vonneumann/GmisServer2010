Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'获取项目组长的员工ID
Public Class ImplFindTeamHeader
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction


    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '定义项目组用户对象引用
    Private Staff As Staff

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

        '实例化项目组用户对象
        Staff = New Staff(conn, ts)

        '实例化角色用户对象
        Role = New Role(conn, ts)

    End Sub

    Public Function UseTools(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '1、获取当前任务的项目组
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id='RegisterTeam'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '异常处理  
        If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            Throw wfErr
        End If

        Dim tmpTeam As String = dsTempTaskAttendee.Tables(0).Rows(0).Item("attend_person")

        '2、在项目组员工表获取参与人（项目组）指定的所有员工
        strSql = "{team_name=" & "'" & tmpTeam & "'" & "}"
        Dim dsTempTeamStaff As DataSet = Staff.FetchStaff(strSql)

        '如果找不到该项目组，则抛出提交无效错误
        If dsTempTeamStaff.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoTeamStaffErr()
            Throw wfErr
            Exit Function
        End If

        '3、对于每个员工
        Dim i, j, k As Integer
        Dim tmpTeamStaff As String
        Dim dsTempRole As DataSet

        Dim bDone As Boolean

        For i = 0 To dsTempTeamStaff.Tables(0).Rows.Count - 1

            '在员工角色表获取员工的所有角色集合
            tmpTeamStaff = dsTempTeamStaff.Tables(0).Rows(i).Item("staff_name")
            dsTempRole = Role.GetStaffRole("%", tmpTeamStaff)

            '如果员工的角色集合包含项目组组长角色；
            ' ' 将本项目角色为项目组长的参与人均修改为员工；
            '        返回
            For j = 0 To dsTempRole.Tables(0).Rows.Count - 1

                '如果员工的角色集合包含项目组组长角色
                If dsTempRole.Tables(0).Rows(j).Item("role_id") = "23" Then
                    Dim tmpHeader As String = dsTempRole.Tables(0).Rows(j).Item("staff_name")
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='AssignProjectManager'" & " and role_id='23'}"
                    dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    '设置标志位,证明已找到组长
                    bDone = True

                    '将本项目角色为项目组长的参与人均修改为员工；
                    For k = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                        dsTempTaskAttendee.Tables(0).Rows(k).Item("attend_person") = tmpHeader
                    Next
                    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
                    Exit Function
                End If
            Next
        Next

        '如果没有,抛出未找到项目组长错误
        If bDone = False Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoTeamHeaderErr()
            Throw wfErr
            Exit Function
        End If

    End Function

End Class
