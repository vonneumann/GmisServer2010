'������Ŀ����A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetManagerA
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
        strSql = "select manager_A,manager_B from QueryProjectInfo where ProjectCode='" & projectID & "'"
        Dim dsProjectInfo As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        '��ȡ��Ŀ����A
        Dim strManagerA As String = IIf(IsDBNull(dsProjectInfo.Tables(0).Rows(0).Item("manager_A")), "", dsProjectInfo.Tables(0).Rows(0).Item("manager_A"))
        '��ȡ��Ŀ����B
        Dim strManagerB As String = IIf(IsDBNull(dsProjectInfo.Tables(0).Rows(0).Item("manager_B")), "", dsProjectInfo.Tables(0).Rows(0).Item("manager_B"))

        '����ɫΪ��Ŀ����A�Ҳ�����Ϊ�յļ�¼������Ŀ����A
        strSql = "{project_code='" & projectID & "' and role_id='24' and isnull(attend_person,'')=''}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        Dim i As Integer
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strManagerA
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

        strSql = "{project_code='" & projectID & "' and role_id='25' and isnull(attend_person,'')=''}"
        dsAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(i).Item("attend_person") = strManagerB
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function

End Class
