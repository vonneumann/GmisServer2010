
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplRecordSignatureCheck
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '������Ϣ��������
    Private WfProjectMessages As WfProjectMessages

    '����ͨ�ò�ѯ��������
    Private CommonQuery As CommonQuery

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

        'ʵ������Ϣ����
        WfProjectMessages = New WfProjectMessages(conn, ts)

        'ʵ����ͨ�ò�ѯ����
        CommonQuery = New CommonQuery(conn, ts)



    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        If userID <> "���º�" Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowMustRecordSignatureSubmitor()
            Throw wfErr
        End If
    End Function

End Class
