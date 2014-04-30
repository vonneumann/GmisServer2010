
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplAddValuator
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
        '��ȡ����Ŀ���ʲ�����ʦ
        Dim strSql As String
        Dim i, j As Integer
        Dim dsTempValuator, dsTempAttend, dsTempDeleteAttend, dsTempRoleTemplate, dsTempTaskMessages, dsTempGuarantyName As DataSet
        Dim iAttendCount, iValuatorCount As Integer
        Dim tmpValuate_person, tmpValuateGuarantyType, tmpValuateGuarantyName, tmpWorkflowID As String
        Dim newRow As DataRow
        Dim itemType As New Item(conn, ts)
        strSql = "SELECT DISTINCT evaluate_person" & _
                 " FROM opposite_guarantee" & _
                 " WHERE evaluate_person is not null" & _
                 " and project_code =" & "'" & projectID & "'"
        dsTempValuator = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsTempValuator.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempValuator.Tables(0))
            Throw wfErr
        End If

        iValuatorCount = dsTempValuator.Tables(0).Rows.Count

        '�Ȼ�ȡ��ǰ�����ģ��ID(��Ϊ����ģ���ʵ��ʱ������)
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id=" & "'" & taskID & "'" & "}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '�쳣����  
        If dsTempDeleteAttend.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowNoRecordkErr(dsTempDeleteAttend.Tables(0))
            Throw wfErr
        End If

        tmpWorkflowID = dsTempDeleteAttend.Tables(0).Rows(0).Item("workflow_id")

        '��ɾ�������ɫ����ԭ���ʲ�����ʦ��������һ�η��������ʦ��
        'strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'" & "}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'" & " and workflow_id='" & tmpWorkflowID & "'}"
        dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        For i = 0 To dsTempDeleteAttend.Tables(0).Rows.Count - 1
            dsTempDeleteAttend.Tables(0).Rows(i).Delete()
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        ''AssignValuator_Update
        ''�������������ʦ�ύ��ȥ��Ȼ���������·�������ʩ�������·�������ʦ������������������ʦ����ǰ�����������ʦ�����¼Ҫɾ����
        'If taskID = "AssignValuator_Update" Then
        '    strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'" & " and task_id='" & tmpWorkflowID & "'}"
        '    dsTempDeleteAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)

        '    For i = 0 To dsTempDeleteAttend.Tables(0).Rows.Count - 1
        '        dsTempDeleteAttend.Tables(0).Rows(i).Delete()
        '    Next
        '    WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)
        'End If



        '�ٸ��������ɫģ����ʲ�����ʦ������
        strSql = "{workflow_id=" & "'" & tmpWorkflowID & "'" & " and role_id='34'" & "}"
        Dim WfTaskRoleTemplate As New WfTaskRoleTemplate(conn, ts)
        dsTempRoleTemplate = WfTaskRoleTemplate.GetWfTaskRoleTemplateInfo(strSql)

        For i = 0 To dsTempRoleTemplate.Tables(0).Rows.Count - 1
            newRow = dsTempDeleteAttend.Tables(0).NewRow()
            With newRow
                .Item("workflow_id") = tmpWorkflowID
                .Item("project_code") = projectID    '�������������Ĺ�����ID��Ϊ��Ŀ����
                .Item("task_id") = dsTempRoleTemplate.Tables(0).Rows(i).Item("task_id")
                .Item("role_id") = dsTempRoleTemplate.Tables(0).Rows(i).Item("role_id")
                .Item("attend_person") = ""
            End With
            dsTempDeleteAttend.Tables(0).Rows.Add(newRow)
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempDeleteAttend)

        ' strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34'}"
        strSql = "{project_code=" & "'" & projectID & "'" & " and role_id='34' and workflow_id='" & tmpWorkflowID & "'}"
        dsTempAttend = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        iAttendCount = dsTempAttend.Tables(0).Rows.Count

        '��	����ÿλ�ʲ�����ʦ
        ' ����Ŀ����������˱���role-id=34������n-1�� ,�������˸�Ϊ�ʲ�����ʦ��
        strSql = "evaluate_person desc"
        dsTempValuator.Tables(0).Select("", strSql)
        Dim tmpPreValuator As String
        If iValuatorCount <> 0 Then


            '���ֻ��һ������ʦ�����ÿ����ɫ����Ĳ�������Ϊ������ʦ

            For i = 0 To iAttendCount - 1

                dsTempAttend.Tables(0).Rows(i).Item("attend_person") = dsTempValuator.Tables(0).Rows(0).Item("evaluate_person")

            Next

            '����ж������ʦ����Ϊ����������ʦ����ÿ����ɫ����

            If iValuatorCount >= 2 Then

                For i = 1 To iValuatorCount - 1

                    For j = 0 To dsTempRoleTemplate.Tables(0).Rows.Count - 1
                        newRow = dsTempAttend.Tables(0).NewRow
                        With dsTempRoleTemplate.Tables(0).Rows(j)
                            newRow.Item("project_code") = projectID
                            newRow.Item("workflow_id") = .Item("workflow_id")
                            newRow.Item("task_id") = .Item("task_id")
                            newRow.Item("role_id") = .Item("role_id")
                            newRow.Item("attend_person") = dsTempValuator.Tables(0).Rows(i).Item("evaluate_person")
                        End With
                        dsTempAttend.Tables(0).Rows.Add(newRow)

                    Next

                Next
             
            End If

            WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempAttend)

        Else

            '�׳�δ�����ʲ�����ʦ����
            Dim wfErr As New WorkFlowErr()
            wfErr.ThrowAddValuatorErr()
            Throw wfErr

        End If



    End Function

End Class
