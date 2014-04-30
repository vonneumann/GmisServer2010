Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'�ж��Ƿ���Ŀ����(ValidateProjectFinished),ֻ�е�"�����������"��"�ͷű�֤��"�������,����ת��"��Ŀ����"(ProjectFinished)
Public Class ImplIsProjectFinished
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    'Private WorkFlow As WorkFlow
    'Private TimingServer As TimingServer
    Private WfProjectTaskTransfer As WfProjectTaskTransfer
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'WorkFlow = New WorkFlow(conn, ts)

        'TimingServer = New TimingServer(conn, ts, True, True)

        WfProjectTaskTransfer = New WfProjectTaskTransfer(conn, ts)

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ò�����ת������
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

    '�ж��Ƿ�VaildateUnfreezeGuaranty��ValidateUnfreezeDepositFee����������(project_stauts='P': �������ڽ���) 
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
