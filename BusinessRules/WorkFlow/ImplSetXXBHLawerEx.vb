'������Ŀ����A
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplSetXXBHLawerEx
    Implements IFlowTools

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction

    '��������˶�������
    Private WfProjectTaskAttendee As WfProjectTaskAttendee

    Private ConfernceRoom As ConfernceRoom

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

        Dim CommonQuery As New CommonQuery(conn, ts)

        Dim dsProjectInfo As DataSet
        '��ȡ��Ŀ����A,B
        strSql = "select nowManagerA,nowManagerB from queryProjectInfo where projectCode='" & projectID & "'"
        dsProjectInfo = CommonQuery.GetCommonQueryInfo(strSql)

        '�쳣����  
        If dsProjectInfo.Tables(0).Rows.Count = 0 Then
            Dim wfErr As New WorkFlowErr
            wfErr.ThrowNoRecordkErr(dsProjectInfo.Tables(0))
            Throw wfErr
        End If

        Dim tmpManagerA As String = dsProjectInfo.Tables(0).Rows(0).Item("nowManagerA")

        '2011-5-20 YJF ADD 
        '���÷�����
        '��ȡ��Ŀ�������ڵĲ���
        strSql = "select dept_name from staff where staff_name='" & tmpManagerA & "'"
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


        '���ñ���Ŀ�ķ�����
        strSql = "{project_code='" & projectID & "' and role_id='33'}"
        Dim dsTempTaskAttendee As DataSet = WfProjectTaskAttendee.GetWfProjectTaskAttendeeInfo(strSql)
        For Each drTemp As DataRow In dsTempTaskAttendee.Tables(0).Rows
            drTemp.Item("attend_person") = strPerson
        Next
        WfProjectTaskAttendee.UpdateWfProjectTaskAttendee(dsTempTaskAttendee)
    End Function

End Class
