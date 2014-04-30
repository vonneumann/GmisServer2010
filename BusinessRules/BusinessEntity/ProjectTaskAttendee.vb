Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectTaskAttendee

    Public Const Table_Project_Task_Attendee As String = "project_task_attendee"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectTaskAttendee As SqlDataAdapter

    '�����ѯ����
    Private GetProjectTaskAttendeeInfoCommand As SqlCommand
    Private GetProjectAttendeeInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectTaskAttendee = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectTaskAttendeeInfo("null")

    End Sub

    '��ȡ��Ŀ��������Ϣ��ȫ����
    Public Function GetProjectTaskAttendeeInfo(ByVal strSQL_Condition_ProjectTaskAttendee As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectTaskAttendeeInfoCommand Is Nothing Then

            GetProjectTaskAttendeeInfoCommand = New SqlCommand("GetProjectTaskAttendeeInfo", conn)
            GetProjectTaskAttendeeInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectTaskAttendeeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectTaskAttendee
            .SelectCommand = GetProjectTaskAttendeeInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectTaskAttendeeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectTaskAttendee
            .Fill(tempDs, Table_Project_Task_Attendee)
        End With

        GetProjectTaskAttendeeInfo = tempDs

    End Function

    '��ȡ��Ŀ��������Ϣ(distinct ������Ψһ)
    Public Function GetProjectAttendeeInfo(ByVal projectID As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectAttendeeInfoCommand Is Nothing Then

            GetProjectAttendeeInfoCommand = New SqlCommand("GetProjectAttendeeInfo", conn)
            GetProjectAttendeeInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectAttendeeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectTaskAttendee
            .SelectCommand = GetProjectAttendeeInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectAttendeeInfoCommand.Parameters("@Condition").Value = projectID
            .Fill(tempDs, Table_Project_Task_Attendee)
        End With

        GetProjectAttendeeInfo = tempDs

    End Function


    '������Ŀ��������Ϣ
    Public Function UpdateProjectTaskAttendee(ByVal ProjectTaskAttendeeSet As DataSet)

        If ProjectTaskAttendeeSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If ProjectTaskAttendeeSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectTaskAttendee)

        With dsCommand_ProjectTaskAttendee

            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectTaskAttendeeSet, Table_Project_Task_Attendee)

        End With

        ProjectTaskAttendeeSet.AcceptChanges()



    End Function

    Public Function GetProjectTaskAttendFromRoleID(ByVal projectID As String, ByVal RoleID As String) As String
        Dim ds As DataSet
        Dim dv As DataView
        Dim i, count As Integer
        Dim strAttend As String

        ds = GetProjectAttendeeInfo(projectID)
        dv = ds.Tables(0).DefaultView
        dv.RowFilter = "role_id='" & RoleID & "' and not attend_person=''"
        count = dv.Count
        If count > 0 Then
            strAttend = dv.Item(0).Item("attend_person")
        End If
        Return strAttend.Trim
    End Function

End Class
