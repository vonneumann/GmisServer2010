Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendHandoverMsg
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    '定义角色对象引用
    Private Role As Role

    '定义时间服务对象引用
    Private TimingServer As TimingServer

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化角色对象
        Role = New Role(conn, ts)

        '实例化时间服务对象
        TimingServer = New TimingServer(conn, ts, True, True)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools


        '①	获取初审统计人员的员工ID=14；
        Dim strSql As String = "{role_id='14'}"
        Dim dsTempRoleStaff As DataSet = Role.GetStaffRole(strSql)
        Dim i As Integer
        Dim tmpUserID As String
        For i = 0 To dsTempRoleStaff.Tables(0).Rows.Count - 1
            tmpUserID = Trim(dsTempRoleStaff.Tables(0).Rows(i).Item("staff_name"))
            '②	调用AddMsg（项目ID，“HandoverRegister”，13、员工、确认标志）
            TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpUserID, "13", "N")
        Next
    End Function
End Class
