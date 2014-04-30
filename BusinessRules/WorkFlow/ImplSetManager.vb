Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient


'���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
Public Class ImplSetManager
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

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

        'ʵ������ɫ�û�����
        Role = New Role(conn, ts)


    End Sub

    '���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        '��ȡ����Ŀ�Ĺ��������������Ϣ
        Dim strAttendee As String
        Dim strSql As String = "{project_code=" & "'" & projectID & "'" & " and task_id='" & "ReviewMeetingPlan" & "'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsTempTaskAttendee.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempTaskAttendee.Tables(0))
            Throw wfErr
        End If

        If Not dsTempTaskAttendee Is Nothing Then
            strAttendee = dsTempTaskAttendee.Tables(0).Rows(0).Item("attend_person")
        End If

        '''''''''''''''
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='" & "RecordReviewConclusion" & "'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        Dim i As Integer
        Dim dsTempRole As DataSet
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1

            '���й̶�Ա���Ľ�ɫ��Ա��ID��ӵ���������
            With dsTempTaskAttendee.Tables(0).Rows(i)
                .Item("attend_person") = strAttendee

                Select Case Trim(.Item("role_id"))
                    Case "01" '��������
                        dsTempRole = Role.GetStaffRole("01")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "12" '���Ų���
                        dsTempRole = Role.GetStaffRole("12")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "21" '��������
                        dsTempRole = Role.GetStaffRole("21")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "31" '���ղ���
                        dsTempRole = Role.GetStaffRole("31")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "40" '�ۺϹ�����
                        dsTempRole = Role.GetStaffRole("40")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "41" '������Ա
                        dsTempRole = Role.GetStaffRole("41")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "14" '����ͳ����Ա
                        dsTempRole = Role.GetStaffRole("14")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")
                    Case "27" '��Ŀ�Ǽ�Ա
                        dsTempRole = Role.GetStaffRole("27")

                        '�쳣����  
                        If dsTempRole.Tables(0).Rows.Count = 0 Then
                            Dim wfErr As New WorkFlowErr()
                            wfErr.ThrowNoRecordkErr(dsTempRole.Tables(0))
                            Throw wfErr
                        End If

                        .Item("attend_person") = dsTempRole.Tables(0).Rows(0).Item("staff_name")

                End Select


            End With
        Next

        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

    End Function

End Class
