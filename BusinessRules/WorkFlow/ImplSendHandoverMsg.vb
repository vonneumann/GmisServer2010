Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendHandoverMsg
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '�����ɫ��������
    Private Role As Role

    '����ʱ������������
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

        'ʵ������ɫ����
        Role = New Role(conn, ts)

        'ʵ����ʱ��������
        TimingServer = New TimingServer(conn, ts, True, True)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools


        '��	��ȡ����ͳ����Ա��Ա��ID=14��
        Dim strSql As String = "{role_id='14'}"
        Dim dsTempRoleStaff As DataSet = Role.GetStaffRole(strSql)
        Dim i As Integer
        Dim tmpUserID As String
        For i = 0 To dsTempRoleStaff.Tables(0).Rows.Count - 1
            tmpUserID = Trim(dsTempRoleStaff.Tables(0).Rows(i).Item("staff_name"))
            '��	����AddMsg����ĿID����HandoverRegister����13��Ա����ȷ�ϱ�־��
            TimingServer.AddMsg(workFlowID, projectID, "HandoverRegister", tmpUserID, "13", "N")
        Next
    End Function
End Class
