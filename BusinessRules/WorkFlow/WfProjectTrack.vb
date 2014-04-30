Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WfProjectTrack

    Public Const Table_Project_Track As String = "project_track"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WfProjectTrack As SqlDataAdapter

    '�����ѯ����
    Private GetWfProjectTrackInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WfProjectTrack = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWfProjectTrackInfo("null")
    End Sub

    '��ȡ��������¼�����Ϣ
    Public Function GetWfProjectTrackInfo(ByVal strSQL_Condition_WfProjectTrack As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWfProjectTrackInfoCommand Is Nothing Then

            GetWfProjectTrackInfoCommand = New SqlCommand("GetWfProjectTrackInfo", conn)
            GetWfProjectTrackInfoCommand.CommandType = CommandType.StoredProcedure
            GetWfProjectTrackInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WfProjectTrack
            .SelectCommand = GetWfProjectTrackInfoCommand
            .SelectCommand.Transaction = ts
            GetWfProjectTrackInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WfProjectTrack
            .Fill(tempDs, Table_Project_Track)
        End With

        Return tempDs
      
    End Function

    '���¹�������¼�����Ϣ
    Public Function UpdateWfProjectTrack(ByVal WfProjectTrackSet As DataSet)

        If WfProjectTrackSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WfProjectTrackSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WfProjectTrack)

        With dsCommand_WfProjectTrack
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(WfProjectTrackSet, Table_Project_Track)

            WfProjectTrackSet.AcceptChanges()
        End With
       
    End Function

End Class
