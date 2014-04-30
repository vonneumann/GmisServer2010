'����������ԱΪ��Ŀ����A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetAppAsManagerA
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


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        strSql = "{project_code='" & projectID & "' and task_id='Application'}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsAttend.Tables(0))
            Throw wfErr
        End If

        '��ȡ��Ŀ����A
        Dim strManagerA As String = IIf(IsDBNull(dsAttend.Tables(0).Rows(0).Item("attend_person")), "", dsAttend.Tables(0).Rows(0).Item("attend_person"))

        '����ɫΪ��Ŀ����A�Ҳ�����Ϊ�յļ�¼������Ŀ����A
        strSql = "{project_code='" & projectID & "' and role_id='24' and isnull(attend_person,'')=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strManagerA
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function
End Class
