'设置项目经理A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetXXBHLawerEx
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义参与人对象引用
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private ConfernceRoom As ConfernceRoom

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

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String

        Dim CommonQuery As New CommonQuery(conn, ts)

        Dim dsProjectInfo As DataSet
        '获取项目经理A,B
        strSql = "select nowManagerA,nowManagerB from queryProjectInfo where projectCode='" & projectID & "'"
        dsProjectInfo = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        Dim tmpManagerA As String = dsProjectInfo.Tables(0).Rows(0).Item("nowManagerA")

        '2011-5-20 YJF ADD 
        '设置法务经理
        '获取项目经理所在的部门
        strSql = "select dept_name from staff where staff_name='" & tmpManagerA & "'"
        Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        strSql = "select staff_name from staff where  isnull(unchain_department_list,'') like '%" & strDeptName & "%'"
        Dim dsTemp2 As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        Dim strPerson As String
        If dsTemp2.Tables(0).Rows.Count <> 0 Then
            strPerson = dsTemp2.Tables(0).Rows(0).Item("staff_name")
        End If


        '设置本项目的法务经理
        strSql = "{project_code='" & projectID & "' and role_id='33'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
            drTemp.Item("attend_person") = strPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
    End Function

End Class
