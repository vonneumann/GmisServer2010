Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplLoanMsg
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction


    '���嶨ʱ�����������
    Private TimingServer As TimingServer

    Private WorkFlow As WorkFlow


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������ʱ�����������
        TimingServer = New TimingServer(conn, ts, True, True)

        WorkFlow = New WorkFlow(conn, ts)

    End Sub

    'ǩ���ſ�֪ͨ��Ӧ֪ͨ���ɰ���ſ����Ϣ�����ɺͻ�ƣ�С����ģ�
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ����
        Dim tmpReciever As String

        ''����ϵͳֻ��һ������,����(ʹ��getTaskActor(roleid ,branch)������ȡ)
        'tmpReciever = WorkFlow.getTaskActor("41")

        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpReciever, "25", "N")

        '2009-10-15 yjf add ֻ��ί����Ŀ����Ҫ���ͷſ���Ϣ
        If workFlowID = "03" Or workFlowID = "05" Then

            '2009��06��12 yjf edit Ϊÿ�����ɷ��Ͱ���ſ���Ϣ������֪���ĸ����к�֧�зſ�
            Dim objCommonQuery As New CommonQuery(conn, ts)
            Dim dsStaff As DataSet = objCommonQuery.GetCommonQueryInfo("select staff_name from staff_role where role_id in ('41','43','45','46')")

            '��ȡ����ĿǩԼ���У�֧��
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
                tmpDr.Item("message_content") = dsBank.Tables(0).Rows(0).Item("EnterpriseName") & "����ſ�����" & " " & _
                                                "����:" & dsBank.Tables(0).Rows(0).Item("sign_bank_name") & " " & _
                                                "֧��:" & dsBank.Tables(0).Rows(0).Item("sign_bank_branch_name") & _
                                                "���:" & dsBank.Tables(0).Rows(0).Item("sign_sum") & "��"
                dsMessage.Tables(0).Rows.Add(tmpDr)
            Next

            objWfProjectMessage.UpdateWfProjectMessages(dsMessage)
            dsMessage.AcceptChanges()
        End If

    End Function

End Class
