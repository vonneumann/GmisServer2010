Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'��ȡ��Ŀ�鳤��Ա��ID
Public Class ImplFindTeamHeader
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction


    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '������Ŀ���û���������
    Private Staff As Staff

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

        'ʵ������Ŀ���û�����
        Staff = New Staff(conn, ts)

        'ʵ������ɫ�û�����
        Role = New Role(conn, ts)

    End Sub

    Public Function UseTools(ByVal workflowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools
        '1����ȡ��ǰ�������Ŀ��
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id='RegisterTeam'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            Throw wfErr
        End If

        Dim tmpTeam As String = dsTempTaskAttendee.Tables(0).Rows(0).Item("attend_person")

        '2������Ŀ��Ա�����ȡ�����ˣ���Ŀ�飩ָ��������Ա��
        strSql = "{team_name=" & "'" & tmpTeam & "'" & "}"
        Dim dsTempTeamStaff As DataSet = Staff.FetchStaff(strSql)

        '����Ҳ�������Ŀ�飬���׳��ύ��Ч����
        If dsTempTeamStaff.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoTeamStaffErr()
            Throw wfErr
            Exit Function
        End If

        '3������ÿ��Ա��
        Dim i, j, k As Integer
        Dim tmpTeamStaff As String
        Dim dsTempRole As DataSet

        Dim bDone As Boolean

        For i = 0 To dsTempTeamStaff.Tables(0).Rows.Count - 1

            '��Ա����ɫ���ȡԱ�������н�ɫ����
            tmpTeamStaff = dsTempTeamStaff.Tables(0).Rows(i).Item("staff_name")
            dsTempRole = Role.GetStaffRole("%", tmpTeamStaff)

            '���Ա���Ľ�ɫ���ϰ�����Ŀ���鳤��ɫ��
            ' ' ������Ŀ��ɫΪ��Ŀ�鳤�Ĳ����˾��޸�ΪԱ����
            '        ����
            For j = 0 To dsTempRole.Tables(0).Rows.Count - 1

                '���Ա���Ľ�ɫ���ϰ�����Ŀ���鳤��ɫ
                If dsTempRole.Tables(0).Rows(j).Item("role_id") = "23" Then
                    Dim tmpHeader As String = dsTempRole.Tables(0).Rows(j).Item("staff_name")
                    strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='AssignProjectManager'" & " and role_id='23'}"
                    dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

                    '���ñ�־λ,֤�����ҵ��鳤
                    bDone = True

                    '������Ŀ��ɫΪ��Ŀ�鳤�Ĳ����˾��޸�ΪԱ����
                    For k = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
                        dsTempTaskAttendee.Tables(0).Rows(k).Item("attend_person") = tmpHeader
                    Next
                    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
                    Exit Function
                End If
            Next
        Next

        '���û��,�׳�δ�ҵ���Ŀ�鳤����
        If bDone = False Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoTeamHeaderErr()
            Throw wfErr
            Exit Function
        End If

    End Function

End Class
