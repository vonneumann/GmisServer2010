Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplLoanMsg
    Implements IFlowTools


    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction


    '定义定时服务对象引用
    Private TimingServer As TimingServer

    Private WorkFlow As WorkFlow


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '实例化定时服务对象引用
        TimingServer = New TimingServer(conn, ts, True, True)

        WorkFlow = New WorkFlow(conn, ts)

    End Sub

    '签发放款通知后应通知出纳办理放款；发消息给出纳和会计（小额、技改）
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获取出纳
        Dim tmpReciever As String

        ''假设系统只有一个出纳,否则(使用getTaskActor(roleid ,branch)方法获取)
        'tmpReciever = WorkFlow.getTaskActor("41")

        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpReciever, "25", "N")

        '2009-10-15 yjf add 只有委贷项目才需要发送放款消息
        If workFlowID = "03" Or workFlowID = "05" Then

            '2009－06－12 yjf edit 为每个出纳发送办理放款消息，并告知由哪个银行和支行放款
            Dim objCommonQuery As New CommonQuery(conn, ts)
            Dim dsStaff As DataSet = objCommonQuery.GetCommonQueryInfo("select staff_name from staff_role where role_id in ('41','43','45','46')")

            '获取本项目签约银行，支行
            Dim dsBank As DataSet = objCommonQuery.GetCommonQueryInfo("select EnterpriseName,sign_sum,sign_bank_name,sign_bank_branch_name from viewProjectInfo where ProjectCode='" & projectID & "'")
            Dim i As Integer
            Dim tmpDr As DataRow
            Dim dsMessage As DataSet
            Dim objWfProjectMessage As New WfProjectMessages(conn, ts)
            dsMessage = objWfProjectMessage.GetWfProjectMessagesInfo("")
            For i = 0 To dsStaff.Tables(0).Rows.Count - 1
                tmpDr = dsMessage.Tables(0).NewRow
                tmpDr.Item("project_code") = projectID
                tmpDr.Item("accepter") = dsStaff.Tables(0).Rows(i).Item("staff_name")
                tmpDr.Item("send_time") = Now
                tmpDr.Item("is_affirmed") = "N"
                tmpDr.Item("message_content") = dsBank.Tables(0).Rows(0).Item("EnterpriseName") & "办理放款手续" & " " & _
                                                "银行:" & dsBank.Tables(0).Rows(0).Item("sign_bank_name") & " " & _
                                                "支行:" & dsBank.Tables(0).Rows(0).Item("sign_bank_branch_name") & _
                                                "金额:" & dsBank.Tables(0).Rows(0).Item("sign_sum") & "万"
                dsMessage.Tables(0).Rows.Add(tmpDr)
            Next

            objWfProjectMessage.UpdateWfProjectMessages(dsMessage)
            dsMessage.AcceptChanges()
        End If

    End Function

End Class
