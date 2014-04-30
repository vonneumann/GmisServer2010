'������Ŀ����A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetLawer
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

        '��ȡ��ǰ������ͬ���û����ڵĲ���
        strSql = "select dept_name from staff where staff_name='" & userID & "'"
        Dim dsTemp As DataSet = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsTemp.Tables(0))
            Throw wfErr
        End If

        Dim strDeptName As String = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("dept_name")), "", dsTemp.Tables(0).Rows(0).Item("dept_name"))

        '��ȡ�ò��ŵĺ�ͬ�����Ա
        strSql = "select staff_name from staff_role where role_id='39'"
        dsTemp = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsTemp.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoStaffRole()
            Throw wfErr
        End If

        Dim i, j As Integer
        Dim strStaff, strConsigner As String
        Dim dsTemp2, dsTemp3 As DataSet
        Dim isFound As Boolean
        For i = 0 To dsTemp.Tables(0).Rows.Count - 1
            strStaff = dsTemp.Tables(0).Rows(i).Item("staff_name")
            strSql = "select staff_name from staff where staff_name='" & strStaff & "' and dept_name='" & strDeptName & "'"
            dsTemp2 = CommonQuery.GetCommonQueryInfo(strSql)
            If dsTemp2.Tables(0).Rows.Count <> 0 Then
                isFound = True
                '�ж��Ƿ�������ί�У����������ί���˴���
                strSql = "select * from staff_role where role_id='39' and staff_name='" & strStaff & "'"
                dsTemp3 = CommonQuery.GetCommonQueryInfo(strSql)
                If dsTemp3.Tables(0).Rows.Count <> 0 Then
                    strConsigner = Trim(IIf(IsDBNull(dsTemp3.Tables(0).Rows(0).Item("consigner")), "", dsTemp3.Tables(0).Rows(0).Item("consigner")))
                    If strConsigner <> "" Then
                        strStaff = strConsigner
                    End If
                End If
                Exit For
            End If
        Next

        '�쳣����  
        If isFound = False Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoStaffRole()
            Throw wfErr
        End If


        '���ñ���Ŀ�ĺ�ͬ�����Ա
        strSql = "{project_code='" & projectID & "' and role_id='39' and isnull(attend_person,'')=''}"
        Dim dsAttend As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For j = 0 To dsAttend.Tables(0).Rows.Count - 1
            dsAttend.Tables(0).Rows(j).Item("attend_person") = strStaff
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsAttend)

    End Function

End Class
