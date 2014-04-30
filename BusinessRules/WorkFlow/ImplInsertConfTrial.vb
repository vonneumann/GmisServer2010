Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'����Ŀ�����ύʱ����conference_trial�����һ����¼

Public Class ImplInsertConfTrial
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    Private conferenceTrial As ConfTrial

    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        conferenceTrial = New ConfTrial(conn, ts)

    End Sub

    Public Function UseTools(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        Dim strSql As String
        Dim dsConferenceTrial As DataSet
        Dim dr As DataRow
        Dim i, count, times As Integer
        Dim status As Boolean

        strSql = "{project_code='" & projectID & "' order by trial_times DESC}"
        dsConferenceTrial = conferenceTrial.GetConfTrialInfo(strSql, "null")
        count = dsConferenceTrial.Tables(0).Rows.Count



        If workflowID = "10" Then

            '���ö�����±����ļ�¼Ա
            strSql = "select record_person from TConferenceRoom where conference_address='���������'"
            Dim CommonQuery As New CommonQuery(conn, ts)
            Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

            '�쳣����  
            If dsTemp.Tables(0).Rows.Count = 0 Then
                Dim wfErr As New WorkFlowErr
                wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
                Throw wfErr
            End If

            Dim strPerson As String = dsTemp.Tables(0).Rows(0).Item("record_person")
            Dim WfProjectTaskAttendee As New WfProjectTaskAttendee(conn, ts)

            strSql = "{project_code='" & projectID & "' and role_id='33'}"
            Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
            For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
                drTemp.Item("attend_person") = strPerson
            Next
            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

        End If

        If count > 0 Then
            times = CInt(dsConferenceTrial.Tables(0).Rows(0).Item("trial_times"))
            status = IIf(dsConferenceTrial.Tables(0).Rows(0).Item("status") Is System.DBNull.Value, False, dsConferenceTrial.Tables(0).Rows(0).Item("status"))
            If status Then
                dr = dsConferenceTrial.Tables(0).NewRow
                With dr
                    .Item("project_code") = projectID
                    .Item("trial_times") = times + 1
                    .Item("status") = 0
                    .Item("create_person") = userID
                End With
                dsConferenceTrial.Tables(0).Rows.Add(dr)
            End If
        Else
            dr = dsConferenceTrial.Tables(0).NewRow
            With dr
                .Item("project_code") = projectID
                .Item("trial_times") = 1
                .Item("status") = 0
                .Item("create_person") = userID
            End With
            dsConferenceTrial.Tables(0).Rows.Add(dr)
        End If
        conferenceTrial.UpdateConfTrial(dsConferenceTrial)
    End Function

End Class
