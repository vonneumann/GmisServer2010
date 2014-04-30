Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSendEndValuationMsg
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���嶨ʱ�����������
    Private WfProjectTimingTask As WfProjectTimingTask

    '��������˶�������
    Private ProjectTaskAttendee As ProjectTaskAttendee

    '������Ϣ��������
    Private WfProjectMessages As WfProjectMessages


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        'ʵ������ʱ�������
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)


        'ʵ���������˶���
        ProjectTaskAttendee = New ProjectTaskAttendee(conn, ts)


        'ʵ������Ϣ����
        WfProjectMessages = New WfProjectMessages(conn, ts)


    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ��Ŀ����
        Dim strSql As String
        Dim i As Integer
        Dim dsTempAttend, dsTempTaskMessages As DataSet
        Dim newRow As DataRow
        Dim tmpManager As String
        Dim drManager() As DataRow
        dsTempAttend = ProjectTaskAttendee.GetProjectAttendeeInfo(projectID)
        drManager = dsTempAttend.Tables(0).Select("role_id in ('24','25')")
        dsTempTaskMessages = WfProjectMessages.GetWfProjectMessagesInfo("null")

        '��ȡ��Ŀ����ҵ����
        strSql = "{project_code=" & "'" & projectID & "'" & "}"
        Dim Project As New Project(conn, ts)
        Dim dsProject As DataSet = Project.GetProjectInfo(strSql)

        '�쳣����  
        If dsProject.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsProject.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationCode As String = dsProject.Tables(0).Rows(0).Item("corporation_code")
        strSql = "{corporation_code=" & "'" & tmpCorporationCode & "'" & "}"
        Dim corporationAccess As New corporationAccess(conn, ts)
        Dim dsCorporation As DataSet = corporationAccess.GetcorporationInfo(strSql, "null")

        '�쳣����  
        If dsCorporation.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsCorporation.Tables(0))
            Throw wfErr
        End If

        Dim tmpCorporationName As String = Trim(dsCorporation.Tables(0).Rows(0).Item("corporation_name"))

        For i = 0 To drManager.Length - 1
            tmpManager = drManager(i).Item("attend_person")
            newRow = dsTempTaskMessages.Tables(0).NewRow
            With newRow
                .Item("project_code") = projectID
                .Item("message_content") = userID & " " & tmpCorporationName & "��Ŀ" & "�ʲ��������"
                .Item("accepter") = tmpManager
                .Item("send_time") = Now
                .Item("is_affirmed") = "N"
            End With
            dsTempTaskMessages.Tables(0).Rows.Add(newRow)
        Next
        WfProjectMessages.UpdateWfProjectMessages(dsTempTaskMessages)

    End Function
End Class
