Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplAllowAccept
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

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

        project = New Project(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim corporationCode As String
        Dim strSql As String

        '获取该项目的企业编码
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim dsTemp As DataSet = project.GetProjectInfo(strSql)

        '异常处理  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If
        '修改project表中的is_accept为0:允许申请
        dsTemp.Tables("project").Rows(0).Item("is_accept") = 0

        project.UpdateProject(dsTemp)

    End Function

End Class
