Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


'���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
Public Class ImplSendMakeContractForBankMsg
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '�����ɫ�û���������
    Private Role As Role


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        'ʵ������ɫ�û�����
        Role = New Role(conn, ts)


    End Sub

    '���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ��Ŀ����A,B
        Dim tmpManagerA, tmpCorporationName, tmpMessage, tmpApplyBank, tmpApplyBranchBank As String
        Dim strsql As String = "select manager_A ,EnterpriseName,ApplyBank,ApplyBranchBank from viewProjectInfo where ProjectCode='" & projectID & "'"

        Dim CommonQuery As New CommonQuery(conn, ts)
        Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strsql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        tmpManagerA = dsProjectInfo.Tables(0).Rows(0).Item("manager_A")
        tmpCorporationName = dsProjectInfo.Tables(0).Rows(0).Item("EnterpriseName")
        tmpApplyBank = dsProjectInfo.Tables(0).Rows(0).Item("ApplyBank")
        tmpApplyBranchBank = dsProjectInfo.Tables(0).Rows(0).Item("ApplyBranchBank")


        Dim TimingServer As New TimingServer(conn, ts, True, True)
        'TimingServer.AddMsg(workFlowID, projectID, taskID, tmpManagerA, "28", "N")
        tmpMessage = projectID & " ����ȷ��Ϊ:" & tmpApplyBank & " " & tmpApplyBranchBank
        TimingServer.AddMsgContent(workFlowID, projectID, taskID, tmpManagerA, tmpMessage, "N")

    End Function

End Class
