Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class CorpDefect

    Public Const Table_Corporation_Defect_Record As String = "corporation_defect_record"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_CorpDefect As SqlDataAdapter

    '�����ѯ����
    Private GetCorpDefectInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_CorpDefect = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetCorpDefectInfo("null")
    End Sub

    '��ȡ��ҵ�۵���Ϣ
    Public Function GetCorpDefectInfo(ByVal strSQL_Condition_CorpDefect As String) As DataSet

        Dim tempDs As New DataSet()

        If GetCorpDefectInfoCommand Is Nothing Then

            GetCorpDefectInfoCommand = New SqlCommand("GetCorpDefectInfo", conn)
            GetCorpDefectInfoCommand.CommandType = CommandType.StoredProcedure
            GetCorpDefectInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_CorpDefect
            .SelectCommand = GetCorpDefectInfoCommand
            .SelectCommand.Transaction = ts
            GetCorpDefectInfoCommand.Parameters("@Condition").Value = strSQL_Condition_CorpDefect
            .Fill(tempDs, Table_Corporation_Defect_Record)
        End With

        GetCorpDefectInfo = tempDs

    End Function

    '������ҵ�۵���Ϣ
    Public Function UpdateCorpDefect(ByVal CorpDefectSet As DataSet)

        If CorpDefectSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If CorpDefectSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_CorpDefect)

        With dsCommand_CorpDefect
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(CorpDefectSet, Table_Corporation_Defect_Record)

        End With

        CorpDefectSet.AcceptChanges()


    End Function

End Class
