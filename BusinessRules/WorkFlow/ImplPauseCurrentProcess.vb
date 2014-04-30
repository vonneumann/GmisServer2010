'��ͣ��ǰ���ڴ��������(���ݻ�����)
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplPauseCurrentProcess
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

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

        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

       
    End Sub


    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim i As Integer

        '�����16���޸ķ�������ʩ���̣���������׶��޸ķ�������ʩ��Ҫ�����ڽ��е��������
        If workFlowID = "16" Then
            Dim strTemp As String
            Dim strPhase As String

            Dim commQuery As CommonQuery = New CommonQuery(conn, ts)
            strTemp = "select phase from project where project_code='" & projectID & "'"
            Dim ds As DataSet = commQuery.GetCommonQueryInfo(strTemp)
            If Not ds Is Nothing Then
                strPhase = IIf(ds.Tables(0).Rows(0).Item("phase") Is System.DBNull.Value, "", ds.Tables(0).Rows(0).Item("phase"))
                If strPhase = "����" Then
                    Exit Function
                End If
            End If
        End If

        '����������IDΪ13��17,22��������е�����״̬ΪP ����Ϊ
        strSql = "{project_code='" & projectID & "' and workflow_id not in ('13','17','22') and task_status='P'}"
        Dim dsTemp As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            dsTemp.Tables(0).Rows(i).Item("task_status") = "C"
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTemp)

    End Function
End Class
