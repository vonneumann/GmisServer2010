'设置修改评审会记录员
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetRecordPerson
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

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

        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        'strSql = "select manager_A from viewProjectInfo where ProjectCode='" & projectID & "'"
        strSql = "{project_code='" & projectID & "' and role_id='51'}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '获取原记录评审会记录员
        Dim strRecordPerson As String
        If dsAttend.Tables(0).Rows.Count > 0 Then
            strRecordPerson = dsAttend.Tables(0).Rows(0).Item("attend_person")
        End If

        '将角色为评审会记录员且参与人为空的记录填入原记录评审会记录员
        strSql = "{project_code='" & projectID & "' and role_id='51' and isnull(attend_person,'')=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strRecordPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function

End Class
