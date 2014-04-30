Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class IntentLetter

    Public Const Table_IntentLetter As String = "intent_letter"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_IntentLetter As SqlDataAdapter

    '�����ѯ����
    Private GetIntentLetterInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_IntentLetter = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetIntentLetterInfo("null")
    End Sub

    '��ȡ��������Ϣ
    Public Function GetIntentLetterInfo(ByVal strSQL_Condition_IntentLetter As String) As DataSet

        Dim tempDs As New DataSet()

        If GetIntentLetterInfoCommand Is Nothing Then

            GetIntentLetterInfoCommand = New SqlCommand("GetIntentLetterInfo", conn)
            GetIntentLetterInfoCommand.CommandType = CommandType.StoredProcedure
            GetIntentLetterInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_IntentLetter
            .SelectCommand = GetIntentLetterInfoCommand
            .SelectCommand.Transaction = ts
            GetIntentLetterInfoCommand.Parameters("@Condition").Value = strSQL_Condition_IntentLetter
            .Fill(tempDs, Table_IntentLetter)
        End With

        Return tempDs

    End Function

    '������������Ϣ
    Public Function UpdateIntentLetter(ByVal IntentLetterSet As DataSet)

        If IntentLetterSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If IntentLetterSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_IntentLetter)

        With dsCommand_IntentLetter
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(IntentLetterSet, Table_IntentLetter)

        End With

        IntentLetterSet.AcceptChanges()

    End Function

End Class
