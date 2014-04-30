Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Conference
    Public Const Table_Conference As String = "conference"
    Public Const Table_Conference_Committeeman As String = "conference_committeeman"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Conference As SqlDataAdapter
    Private dsCommand_Conference_Committeeman As SqlDataAdapter

    '�����ѯ����
    Private GetConferenceInfoCommand As SqlCommand
    Private GetConferenceCommitteemanInfoCommand As SqlCommand
    Private GetMaxConferenceCodeNumCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Conference = New SqlDataAdapter()
        dsCommand_Conference_Committeeman = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetConferenceInfo("null", "null")
    End Sub

    '��ȡ������Ϣ
    Public Function GetConferenceInfo(ByVal strSQL_Condition_Conference As String, ByVal strSQL_Condition_Conference_Committeeman As String) As DataSet

        Dim tempDs As New DataSet()

        If GetConferenceInfoCommand Is Nothing Then

            GetConferenceInfoCommand = New SqlCommand("GetConferenceInfo", conn)
            GetConferenceInfoCommand.CommandType = CommandType.StoredProcedure
            GetConferenceInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Conference
            .SelectCommand = GetConferenceInfoCommand
            .SelectCommand.Transaction = ts
            GetConferenceInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Conference
            .Fill(tempDs, Table_Conference)
        End With

        If GetConferenceCommitteemanInfoCommand Is Nothing Then

            GetConferenceCommitteemanInfoCommand = New SqlCommand("GetConferenceCommitteemanInfo", conn)
            GetConferenceCommitteemanInfoCommand.CommandType = CommandType.StoredProcedure
            GetConferenceCommitteemanInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Conference_Committeeman
            .SelectCommand = GetConferenceCommitteemanInfoCommand
            .SelectCommand.Transaction = ts
            GetConferenceCommitteemanInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Conference_Committeeman
            .Fill(tempDs, Table_Conference_Committeeman)
        End With

        GetConferenceInfo = tempDs

    End Function

    '��ȡ������к�
    Public Function GetMaxConferenceCodeNum() As Integer

        If GetMaxConferenceCodeNumCommand Is Nothing Then

            GetMaxConferenceCodeNumCommand = New SqlCommand("GetMaxConferenceCodeNum", conn)
            GetMaxConferenceCodeNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxConferenceCodeNumCommand.Parameters.Add(New SqlParameter("@maxConferenceCodeNum", SqlDbType.Int))
            GetMaxConferenceCodeNumCommand.Parameters.Item("@maxConferenceCodeNum").Direction = ParameterDirection.Output
            GetMaxConferenceCodeNumCommand.Transaction = ts
        End If

        GetMaxConferenceCodeNumCommand.ExecuteNonQuery()
        GetMaxConferenceCodeNum = GetMaxConferenceCodeNumCommand.Parameters.Item("@maxConferenceCodeNum").Value
    End Function

    '���»�����Ϣ
    Private Function UpdateConference(ByVal ConferenceSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Conference)

        With dsCommand_Conference
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ConferenceSet, Table_Conference)

            ConferenceSet.AcceptChanges()
        End With

    End Function

    '������ί��Ϣ
    Private Function UpdateCommitteeman(ByVal CommitteemanSet As DataSet)

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Conference_Committeeman)

        With dsCommand_Conference_Committeeman
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CommitteemanSet, Table_Conference_Committeeman)

            CommitteemanSet.AcceptChanges()
        End With

    End Function


    '���»���,��ί�����Ϣ
    Public Function UpdateConferenceCommitteeman(ByVal ConferenceCommitteemanSet As DataSet)


        If ConferenceCommitteemanSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ConferenceCommitteemanSet.HasChanges = False Then
            Exit Function
        End If

        'ɾ������
        If IsNothing(ConferenceCommitteemanSet.GetChanges(DataRowState.Deleted)) = False Then
            UpdateCommitteeman(ConferenceCommitteemanSet.GetChanges(DataRowState.Deleted))
            UpdateConference(ConferenceCommitteemanSet.GetChanges(DataRowState.Deleted))
        End If


        '��������
        If IsNothing(ConferenceCommitteemanSet.GetChanges(DataRowState.Added)) = False Then

            ''ͬʱ�����������ϸ��
            'If ConferenceCommitteemanSet.GetChanges(DataRowState.Added).Tables(0).Rows.Count <> 0 And ConferenceCommitteemanSet.GetChanges(DataRowState.Added).Tables(1).Rows.Count <> 0 Then

            '    '��������
            '    Dim i, j As Integer
            '    Dim tmpRowPrimary, tmpRowDetail As DataRow
            '    For i = 0 To ConferenceCommitteemanSet.Tables(0).Rows.Count - 1
            '        tmpRowPrimary = ConferenceCommitteemanSet.Tables(0).Rows(i)
            '        If tmpRowPrimary.RowState = DataRowState.Added Then
            '            '�����ͻ��˵���Ŀ����
            '            Dim projectCode As String = Trim(ConferenceCommitteemanSet.Tables(0).Rows(i).Item("project_code"))
            '            Dim strSql As String = "{project_code=" & "'" & projectCode & "'" & " order by serial_num}"

            '            '��ȡ���ݿ��и���Ŀ�Ļ����¼��Ϊ��ȡ������������룩
            '            Dim dsTemp As DataSet = GetConferenceInfo(strSql, "null")
            '            Dim rowNum As Integer = dsTemp.Tables(0).Rows.Count

            '            '�������ݿ��������������
            '            Dim serialNum As Integer
            '            If rowNum = 0 Then
            '                serialNum = 1
            '            Else
            '                serialNum = dsTemp.Tables(0).Rows(rowNum - 1).Item("serial_num") + 1
            '            End If

            '            '�����ͻ��������������
            '            Dim serialNumTemp As Integer = ConferenceCommitteemanSet.Tables(0).Rows(i).Item("serial_num")

            '            '�ѿͻ��������ڸ����������������������Ϊ���ݿ��е������������
            '            ConferenceCommitteemanSet.Tables(0).Rows(i).Item("serial_num") = serialNum

            '            '�ѿͻ��������ڸ�����������Ļ����������ϸ������������Ϊ���ݿ��е������������
            '            For j = 0 To ConferenceCommitteemanSet.Tables(1).Rows.Count - 1
            '                tmpRowDetail = ConferenceCommitteemanSet.Tables(1).Rows(j)
            '                If tmpRowDetail.RowState = DataRowState.Added Then
            '                    If ConferenceCommitteemanSet.Tables(1).Rows(j).Item("serial_num") = serialNumTemp Then
            '                        ConferenceCommitteemanSet.Tables(1).Rows(j).Item("serial_num") = serialNum
            '                    End If
            '                End If
            '            Next
            '        End If
            '    Next
            'End If

            UpdateConference(ConferenceCommitteemanSet.GetChanges(DataRowState.Added))
            UpdateCommitteeman(ConferenceCommitteemanSet.GetChanges(DataRowState.Added))

        End If

        '�޸Ĳ���
        If IsNothing(ConferenceCommitteemanSet.GetChanges(DataRowState.Modified)) = False Then

            '����ǵ���������Ϣ��ֱ�ӵ����·���
            UpdateConference(ConferenceCommitteemanSet.GetChanges(DataRowState.Modified))
            UpdateCommitteeman(ConferenceCommitteemanSet.GetChanges(DataRowState.Modified))

        End If

        ConferenceCommitteemanSet.AcceptChanges()
    End Function
End Class
