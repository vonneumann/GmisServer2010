'更换还款登记任务的角色为法务经理
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplChgRefundRecordRole
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
        '获取法务经理角色的参与人
        Dim strSql As String
        Dim i As Integer
        Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)
        strSql = "{project_code='" & projectID & "' and role_id='35'}"
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If dsTemp.Tables(0).Rows.Count > 0 Then
            Dim tmpAttend As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("attend_person")), "", dsTemp.Tables(0).Rows(0).Item("attend_person"))

            '将还款登记任务的角色置为'35',参与人置为分配的法务经理
            strSql = "{project_code='" & projectID & "' and task_id='RefundRecord'}"
            dsTemp = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For i = 0 To dsTemp.Tables(0).Rows.Count - 1
                dsTemp.Tables(0).Rows(i).Item("role_id") = "35"
                dsTemp.Tables(0).Rows(i).Item("attend_person") = tmpAttend
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)
        End If

    End Function
End Class
