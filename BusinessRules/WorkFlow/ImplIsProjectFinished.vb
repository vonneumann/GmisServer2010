Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'判断是否项目结束(ValidateProjectFinished),只有当"解除反担保物"和"释放保证金"任务都完成,才流转到"项目结束"(ProjectFinished)
Public Class ImplIsProjectFinished
    Implements IFlowTools

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义事务
    Private ts As SqlTransaction

    'Private WorkFlow As WorkFlow
    'Private TimingServer As TimingServer
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        'WorkFlow = New WorkFlow(conn, ts)

        'TimingServer = New TimingServer(conn, ts, True, True)

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '获得并设置转移条件
        Dim strSql As String
        Dim i As Integer
        Dim dsTempTaskTrans As DataSet

        If Not isProjectFinished(projectID) Then
            strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='ValidateProjectFinished' and next_task='ProjectFinished'}"
            dsTempTaskTrans = WfProjectTaskTransfer.GetWfProjectTaskTransferInfo(strSql)
            If dsTempTaskTrans.Tables(0).Rows.Count > 0 Then
                dsTempTaskTrans.Tables(0).Rows(i).Item("transfer_condition") = ".T."
                WfProjectTaskTransfer.UpdateWfProjectTaskTransfer(dsTempTaskTrans)
            End If
        End If
    End Function

    '判断是否VaildateUnfreezeGuaranty和ValidateUnfreezeDepositFee任务都做完了(project_stauts='P': 任务正在进行) 
    Private Function isProjectFinished(ByVal projectID As String) As Boolean
        Dim strSql As String
        Dim dsTempTask As DataSet

        strSql = "{project_code='" & projectID & "' and (task_id='VaildateUnfreezeGuaranty' or task_id='ValidateUnfreezeDepositFee') and task_status='P'}"
        dsTempTask = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        If Not dsTempTask Is Nothing Then
            If dsTempTask.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


End Class
