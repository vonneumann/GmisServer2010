Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class WorkLog

    Public Const Table_Work_Log As String = "work_log"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_WorkLog As SqlDataAdapter

    '�����ѯ����
    Private GetWorkLogInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_WorkLog = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetWorkLogInfo("null")
    End Sub

    '��ȡ������¼��Ϣ
    Public Function GetWorkLogInfo(ByVal strSQL_Condition_WorkLog As String) As DataSet

        Dim tempDs As New DataSet()

        If GetWorkLogInfoCommand Is Nothing Then

            GetWorkLogInfoCommand = New SqlCommand("GetWorkLogInfo", conn)
            GetWorkLogInfoCommand.CommandType = CommandType.StoredProcedure
            GetWorkLogInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_WorkLog
            .SelectCommand = GetWorkLogInfoCommand
            .SelectCommand.Transaction = ts
            GetWorkLogInfoCommand.Parameters("@Condition").Value = strSQL_Condition_WorkLog
            .Fill(tempDs, Table_Work_Log)
        End With

        Return tempDs


    End Function

    '���¹�����¼��Ϣ
    Public Function UpdateWorkLog(ByVal WorkLogSet As DataSet)

        If WorkLogSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If WorkLogSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_WorkLog)

        With dsCommand_WorkLog
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            '    '��ȡ��¼����
            '    Dim tmpProjectID As String = WorkLogSet.Tables(0).Rows(0).Item("project_code")
            '    Dim strSql As String = "{project_code=" & "'" & tmpProjectID & "'" & " order by serial_num}"
            '    Dim dsTemp As DataSet = GetWorkLogInfo(strSql)
            '    Dim tmpSerialNum As Integer
            '    Dim tmpRowNum As Integer = dsTemp.Tables(0).Rows.Count
            '    If tmpRowNum = 0 Then
            '        '����Ҳ�����¼�����Ŵ�1��ʼ
            '        tmpSerialNum = 1
            '    Else
            '        '   ��ȡ��ǰ��Ŀ����������1
            '        tmpSerialNum = dsTemp.Tables(0).Rows(tmpRowNum - 1).Item("serial_num") + 1
            '    End If

            '    Dim i As Integer
            '    For i = 0 To WorkLogSet.Tables(0).Rows.Count - 1
            '        WorkLogSet.Tables(0).Rows(i).Item("serial_num") = tmpSerialNum
            '        tmpSerialNum = tmpSerialNum + 1
            '    Next

            .Update(WorkLogSet, Table_Work_Log)

            WorkLogSet.AcceptChanges()
        End With


    End Function
End Class
