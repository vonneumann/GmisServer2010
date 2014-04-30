Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Conference
    Public Const Table_Conference As String = "conference"
    Public Const Table_Conference_Committeeman As String = "conference_committeeman"

    '定义全局数据库连接对象
    Private conn As SqlConnection

    '定义全局数据库连接适配器
    Private dsCommand_Conference As SqlDataAdapter
    Private dsCommand_Conference_Committeeman As SqlDataAdapter

    '定义查询命令
    Private GetConferenceInfoCommand As SqlCommand
    Private GetConferenceCommitteemanInfoCommand As SqlCommand
    Private GetMaxConferenceCodeNumCommand As SqlCommand

    '定义事务
    Private ts As SqlTransaction

    '构造函数
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '实例化适配器
        dsCommand_Conference = New SqlDataAdapter()
        dsCommand_Conference_Committeeman = New SqlDataAdapter()

        '打开数据库连接
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '引用外部事务
        ts = trans

        '填充适配器
        GetConferenceInfo("null", "null")
    End Sub

    '获取会议信息
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

    '获取最大序列号
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

    '更新会议信息
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

    '更新评委信息
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


    '更新会议,评委意见信息
    Public Function UpdateConferenceCommitteeman(ByVal ConferenceCommitteemanSet As DataSet)


        If ConferenceCommitteemanSet Is Nothing Then
            Exit Function
        End If

        '如果记录集未发生任何变化，则退出过程
        If ConferenceCommitteemanSet.HasChanges = False Then
            Exit Function
        End If

        '删除操作
        If IsNothing(ConferenceCommitteemanSet.GetChanges(DataRowState.Deleted)) = False Then
            UpdateCommitteeman(ConferenceCommitteemanSet.GetChanges(DataRowState.Deleted))
            UpdateConference(ConferenceCommitteemanSet.GetChanges(DataRowState.Deleted))
        End If


        '新增操作
        If IsNothing(ConferenceCommitteemanSet.GetChanges(DataRowState.Added)) = False Then

            ''同时新增主表和明细表
            'If ConferenceCommitteemanSet.GetChanges(DataRowState.Added).Tables(0).Rows.Count <> 0 And ConferenceCommitteemanSet.GetChanges(DataRowState.Added).Tables(1).Rows.Count <> 0 Then

            '    '批量更新
            '    Dim i, j As Integer
            '    Dim tmpRowPrimary, tmpRowDetail As DataRow
            '    For i = 0 To ConferenceCommitteemanSet.Tables(0).Rows.Count - 1
            '        tmpRowPrimary = ConferenceCommitteemanSet.Tables(0).Rows(i)
            '        If tmpRowPrimary.RowState = DataRowState.Added Then
            '            '读出客户端的项目编码
            '            Dim projectCode As String = Trim(ConferenceCommitteemanSet.Tables(0).Rows(i).Item("project_code"))
            '            Dim strSql As String = "{project_code=" & "'" & projectCode & "'" & " order by serial_num}"

            '            '获取数据库中该项目的会议记录（为获取其最大条数号码）
            '            Dim dsTemp As DataSet = GetConferenceInfo(strSql, "null")
            '            Dim rowNum As Integer = dsTemp.Tables(0).Rows.Count

            '            '读出数据库中最大条数号码
            '            Dim serialNum As Integer
            '            If rowNum = 0 Then
            '                serialNum = 1
            '            Else
            '                serialNum = dsTemp.Tables(0).Rows(rowNum - 1).Item("serial_num") + 1
            '            End If

            '            '读出客户端虚的条数号码
            '            Dim serialNumTemp As Integer = ConferenceCommitteemanSet.Tables(0).Rows(i).Item("serial_num")

            '            '把客户端中属于该虚条数会议的条数号码置为数据库中的最大条数号码
            '            ConferenceCommitteemanSet.Tables(0).Rows(i).Item("serial_num") = serialNum

            '            '把客户端中属于该虚条数会议的会议参与人明细的条数号码置为数据库中的最大条数号码
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

        '修改操作
        If IsNothing(ConferenceCommitteemanSet.GetChanges(DataRowState.Modified)) = False Then

            '如果是单独更改信息，直接调更新方法
            UpdateConference(ConferenceCommitteemanSet.GetChanges(DataRowState.Modified))
            UpdateCommitteeman(ConferenceCommitteemanSet.GetChanges(DataRowState.Modified))

        End If

        ConferenceCommitteemanSet.AcceptChanges()
    End Function
End Class
