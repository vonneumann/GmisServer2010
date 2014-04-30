'�����޸�������¼Ա
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetRecordPerson
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

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

        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        'strSql = "select manager_A from viewProjectInfo where ProjectCode='" & projectID & "'"
        strSql = "{project_code='" & projectID & "' and role_id='51'}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '��ȡԭ��¼������¼Ա
        Dim strRecordPerson As String
        If dsAttend.Tables(0).Rows.Count > 0 Then
            strRecordPerson = dsAttend.Tables(0).Rows(0).Item("attend_person")
        End If

        '����ɫΪ������¼Ա�Ҳ�����Ϊ�յļ�¼����ԭ��¼������¼Ա
        strSql = "{project_code='" & projectID & "' and role_id='51' and isnull(attend_person,'')=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strRecordPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function

End Class
