Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


'把有固定员工的角色的员工ID添加到参与人中
Public Class ImplSendMakeContractMsg
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

        If finishedFlag = "同意" Then
            '获取项目经理A,B
            Dim tmpManagerA As String
            Dim strsql As String = "select manager_A from queryProjectInfo where ProjectCode='" & projectID & "'"

            Dim CommonQuery As New CommonQuery(conn, ts)
            Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strsql)

            '异常处理  
            If dsProjectInfo.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
                Throw wfErr
            End If

            tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("manager_A")

            Dim TimingServer As New TimingServer(conn, ts, True, True)
            'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpManagerA, "28", "N")
            TimingServer.AddMsgContent(workFlowID, projectID, taskID, tmpManagerA, "合同审核通过,请项目经理打印合同!", "N")

        End If
    End Function

End Class
