Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTaskAttendee
    Public Const Table_Project_Task_Attendee As String = "project_task_attendee"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectTaskAttendee As SqlDataAdapter

    '�����ѯ����
    Private GetWfProjectTaskAttendeeInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectTaskAttendee = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectTaskAttendeeInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetWfProjectTaskAttendeeInfo(ByVal strSQL_Condition_WfProjectTaskAttendee As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTaskAttendeeInfoCommand Is Nothing Then

            GetWfProjectTaskAttendeeInfoCommand = New SqlCommand("GetWfProjectTaskAttendeeInfo", conn)
            GetWfProjectTaskAttendeeInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTaskAttendeeInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTaskAttendee
            .SelectCommand = GetWfProjectTaskAttendeeInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTaskAttendeeInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTaskAttendee
            .Fill(tempDs, Table_Project_Task_Attendee)
        End With

        Return tempDs
      
    End Function

    '������Ŀ������Ϣ
    Public Function UpdateWfProjectTaskAttendee(ByVal WfProjectTaskAttendeeSet As DataSet)


        If WfProjectTaskAttendeeSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfProjectTaskAttendeeSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTaskAttendee)

        With dsCommand_WfProjectTaskAttendee
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTaskAttendeeSet, Table_Project_Task_Attendee)

            WfProjectTaskAttendeeSet.AcceptChanges()

        End With
    

    End Function

End Class
