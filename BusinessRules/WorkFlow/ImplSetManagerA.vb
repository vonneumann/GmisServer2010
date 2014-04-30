'设置项目经理A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetManagerA
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
        strSql = "select manager_A,manager_B from QueryProjectInfo where ProjectCode='" & projectID & "'"
        Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        '获取项目经理A
        Dim strManagerA As String = IIf(IsDBNull(dsProjectInfo.Tables(0).Rows(0).Item("manager_A")), "", dsProjectInfo.Tables(0).Rows(0).Item("manager_A"))
        '获取项目经理B
        Dim strManagerB As String = IIf(IsDBNull(dsProjectInfo.Tables(0).Rows(0).Item("manager_B")), "", dsProjectInfo.Tables(0).Rows(0).Item("manager_B"))

        '将角色为项目经理A且参与人为空的记录填入项目经理A
        strSql = "{project_code='" & projectID & "' and role_id='24' and isnull(attend_person,'')=''}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strManagerA
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        strSql = "{project_code='" & projectID & "' and role_id='25' and isnull(attend_person,'')=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strManagerB
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function

End Class
