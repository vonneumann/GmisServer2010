Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplVaildateUnfreezeGuaranty
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private WorkFlow As Workflow

    Private TimingServer As TimingServer

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        Workflow = New WorkFlow(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)
    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '�򵵰�������Ա�ͷ��չ�����������������������Ϣ
        Dim tmpFileManager, tmpMinister As String
        tmpFileManager = WorkFlow.getTaskActor("42")
        tmpMinister = WorkFlow.getTaskActor("31")

        TimingServer.AddMsg(workFlowID, projectID, taskID, tmpFileManager, "23", "N")
        TimingServer.AddMsg(workFlowID, projectID, taskID, tmpMinister, "23", "N")
    End Function
End Class
