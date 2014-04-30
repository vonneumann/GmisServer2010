Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

'ɾ����ʱ����
Public Class ImplDeleteTimingTask
    Implements IFlowTools


    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '���嶨ʱ�����������
    Private WfProjectTimingTask As WfProjectTimingTask

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    '���幤������������
    Private WorkFlow As WorkFlow

    Private TimingServer As TimingServer

    Private workLog As workLog

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

        'ʵ������ʱ�������
        WfProjectTimingTask = New WfProjectTimingTask(conn, ts)

        'ʵ���������˶���
        WfProjectTaskAttendee = New WfProjectTaskAttendee(conn, ts)

        'ʵ��������������
        WorkFlow = New WorkFlow(conn, ts)

        workLog = New WorkLog(conn, ts)

        TimingServer = New TimingServer(conn, ts, True, True)

        CommonQuery = New CommonQuery(conn, ts)

    End Sub

    Public Function UseFlowTools(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal finishedFlag As String, ByVal userID As String) Implements IFlowTools.UseFlowTools

        Dim strSql As String
        Dim dsTempTaskAttendee As DataSet
        Dim i As Integer

        '�ٽ�����Ǽ�����TID=RefundRecord��״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundRecord'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RefundRecord", userID)

        ' '2013-11-30 yjf add ����ȡί����Ϣ����TID=LoanInterestFee��״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='LoanInterestFee'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "LoanInterestFee", userID)


        '��	���ǼǱ�������¼����TID=RecordProjectProcess��״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectProcess", userID)

        '��	����˱�������¼����TID=CheckProjectProcess��״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "CheckProjectProcess", userID)

        '��	��������Ŀ��չ����TID=AppraiseProjectProcess��״̬��Ϊ��F����
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='AppraiseProjectProcess'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "AppraiseProjectProcess", userID)

        '���ǼǱ�����¼����TID=RecordProjectTraceInfo��״̬��Ϊ��F��
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id ='RecordProjectTraceInfo'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)


        '�����˱�����¼����TID=CheckProjectTraceInfo��״̬��Ϊ��F��
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='CheckProjectTraceInfo'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)

        '���Ǽǻ���֤��������TID=RecordRefundCertificate��״̬��Ϊ��F��
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RecordRefundCertificate'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)

        'qxd start
        '���Ǽ���Ŀ������Ϣ����TID=RecordRefundCertificate��״̬��Ϊ��F��
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundDebtInfo_claim'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)

        '���Ǽ�׷�������TID=RecordRefundCertificate��״̬��Ϊ��F��
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id='RefundDebtTrailRecord_claim'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = "F"
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "RecordProjectTraceInfo", userID)
        'end qxd

        '��	�ڶ�ʱ�������ɾ���루ģ��ID����Ŀ���룩ƥ��Ķ�ʱ����Ϊ��T����P�������ж�ʱ����
        strSql = "{project_code=" & "'" & projectID & "'" & " and status in ('T','P')}"
        Dim dsTempTimingTask As DataSet = WfProjectTimingTask.GetWfProjectTimingTaskInfo(strSql)
        For i = 0 To dsTempTimingTask.Tables(0).Rows.Count - 1
            dsTempTimingTask.Tables(0).Rows(i).Delete()
        Next

        WfProjectTimingTask.UpdateWfProjectTimingTask(dsTempTimingTask)

        '���Ǽ�������Ϣ,��¼���ڻ,�ǼǴ�����Ϣ,�ǼǴ��������TID=RecordProjectTraceInfo��״̬��Ϊ����

        '2007-07-12 yjf add
        '���ڷ���ķ��䷨�ﾭ��ҲҪ�ر�
        '�鵵��Ŀ���Ϲر�
        strSql = "{project_code=" & "'" & projectID & "'" & " and task_id in ('SubmissionProjectArchives','OverdueRecord','Overdue_AssignBarrister','OverdueTrailRecord','RefundDebtInfo','RefundDebtTrailRecord')}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For i = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
            dsTempTaskAttendee.Tables(0).Rows(i).Item("task_status") = ""
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
        WorkFlow.AACKMassage(workFlowID, projectID, "OverdueRecord", userID)
        WorkFlow.AACKMassage(workFlowID, projectID, "RefundDebtInfo", userID)
        WorkFlow.AACKMassage(workFlowID, projectID, "OverdueTrailRecord", userID)
        WorkFlow.AACKMassage(workFlowID, projectID, "RefundDebtTrailRecord", userID)

        'sendMesgToManager(workFlowID, projectID)



        '2011-5-20 YJF ADD 
        '����ע����Ա
        '��ȡ��Ŀ�������ڵĲ���
        strSql = "select dept_name from staff where staff_name='" & userID & "'"
        Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        strSql = "select staff_name from staff where  isnull(unchain_department_list,'') like '%" & strDeptName & "%'"
        Dim dsTemp2 As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        Dim strPerson As String
        If dsTemp2.Tables(0).Rows.Count <> 0 Then
            strPerson = dsTemp2.Tables(0).Rows(0).Item("staff_name")
        End If


        '���ñ���Ŀ��ע����Ա
        strSql = "{project_code='" & projectID & "' and role_id='56'}"
        dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
            drTemp.Item("attend_person") = strPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)



        ''2008-5-13 YJF ADD 
        ''����ע���ļ�������
        ''��ȡ��Ŀ�������ڵĲ���
        'strSql = "select dept_name from staff where staff_name='" & userID & "'"
        'Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        ''�쳣����  
        'If dsTemp.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
        '    Throw wfErr
        'End If

        'Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        ''��ȡ����ע���ļ�������
        'strSql = "select staff_name from staff_role where role_id='56'"
        'dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

        ''�쳣����  
        'If dsTemp.Tables(0).Rows.Count = 0 Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoStaffRole()
        '    Throw wfErr
        'End If

        'Dim j As Integer
        'Dim strStaff, strConsigner As String
        'Dim dsTemp2, dsTemp3 As DataSet
        'Dim isFound As Boolean
        'For i = 0 To dsTemp.Tables(0).Rows.Count - 1
        '    strStaff = dsTemp.Tables(0).Rows(i).Item("staff_name")
        '    strSql = "select staff_name from staff where staff_name='" & strStaff & "' and isnull(unchain_department_list,'') like '%" & strDeptName & "%'"
        '    dsTemp2 = CommonQuery.GetCommonQueryInfo(strSql)
        '    If dsTemp2.Tables(0).Rows.Count <> 0 Then
        '        isFound = True
        '        '�ж��Ƿ�������ί�У����������ί���˴���
        '        strSql = "select * from staff_role where role_id='56' and staff_name='" & strStaff & "'"
        '        dsTemp3 = CommonQuery.GetCommonQueryInfo(strSql)
        '        If dsTemp3.Tables(0).Rows.Count <> 0 Then
        '            strConsigner = Trim(IIf(IsDBNull(dsTemp3.Tables(0).Rows(0).Item("consigner")), "", dsTemp3.Tables(0).Rows(0).Item("consigner")))
        '            If strConsigner <> "" Then
        '                strStaff = strConsigner
        '            End If
        '        End If
        '        Exit For
        '    End If
        'Next

        ''�쳣����  
        'If isFound = False Then
        '    Dim wfErr As New WorkFlowErr
        '    wfErr.ThrowNoStaffRole()
        '    Throw wfErr
        'End If


        ''���ñ���Ŀ��ע���ļ�������
        'strSql = "{project_code='" & projectID & "' and role_id='56'}"
        'dsTempTaskAttendee = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        'For j = 0 To dsTempTaskAttendee.Tables(0).Rows.Count - 1
        '    dsTempTaskAttendee.Tables(0).Rows(j).Item("attend_person") = strStaff
        'Next
        'WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)

    End Function

    Private Function sendMesgToManager(ByVal workFlowID As String, ByVal projectID As String)
        Dim strRoleID As String = "31" '���չ�������role_id
        Dim count As Integer
        Dim strAttend As String
        Dim strSql As String
        Dim dsTemp As DataSet

        strSql = "{project_code='" & projectID & "' and role_id='" & strRoleID & "'}"
        dsTemp = workLog.GetWorkLogInfo(strSql)
        If Not dsTemp Is Nothing Then
            count = dsTemp.Tables(0).Rows.Count
            If count > 0 Then
                strAttend = dsTemp.Tables(0).Rows(0).Item("attend_person")
                TimingServer.AddMsg(workFlowID, projectID, "RecordRefundCertificate", strAttend, "27", "N") '27:��Message_template�е�id
            End If
        End If

    End Function
End Class
