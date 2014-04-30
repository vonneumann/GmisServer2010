'设置项目经理A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetLawer
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

        '获取当前制作合同的用户所在的部门
        strSql = "select dept_name from staff where staff_name='" & userID & "'"
        Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        '获取该部门的合同审核人员
        strSql = "select staff_name from staff_role where role_id='39'"
        dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoStaffRole()
            Throw wfErr
        End If

        Dim i, j As Integer
        Dim strStaff, strConsigner As String
        Dim dsTemp2, dsTemp3 As DataSet
        Dim isFound As Boolean
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            strStaff = dsTemp.Tables(0).Rows(i).Item("staff_name")
            strSql = "select staff_name from staff where staff_name='" & strStaff & "' and dept_name='" & strDeptName & "'"
            dsTemp2 = CommonQuery.GetCommonQueryInfo(strSql)
            If dsTemp2.Tables(0).Rows.Count <> 0 Then
                isFound = True
                '判断是否有设置委托，如果有则由委托人处理
                strSql = "select * from staff_role where role_id='39' and staff_name='" & strStaff & "'"
                dsTemp3 = CommonQuery.GetCommonQueryInfo(strSql)
                If dsTemp3.Tables(0).Rows.Count <> 0 Then
                    strConsigner = Trim(IIf(IsDBNull(dsTemp3.Tables(0).Rows(0).Item("consigner")), "", dsTemp3.Tables(0).Rows(0).Item("consigner")))
                    If strConsigner <> "" Then
                        strStaff = strConsigner
                    End If
                End If
                Exit For
            End If
        Next

        '异常处理  
        If isFound = False Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoStaffRole()
            Throw wfErr
        End If


        '设置本项目的合同审核人员
        strSql = "{project_code='" & projectID & "' and role_id='39' and isnull(attend_person,'')=''}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For j = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(j).Item("attend_person") = strStaff
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function

End Class
