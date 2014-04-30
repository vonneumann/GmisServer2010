Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class PawnCredit

    Public Const Table_Pawn_Credit As String = "TPawn_Credit"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Pawn_Credit As SqlDataAdapter


    '�����ѯ����
    Private GetPawnCreditInfoCommand As SqlCommand

    Private GetMaxPawnNumCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Pawn_Credit = New SqlDataAdapter

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetPawnCreditInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetPawnCreditInfo(ByVal strSQL_Condition_Pawn_Credit As String) As DataSet

        Dim tempDs As New DataSet

        If GetPawnCreditInfoCommand Is Nothing Then

            GetPawnCreditInfoCommand = New SqlCommand("GetPawnCreditInfo", conn)
            GetPawnCreditInfoCommand.CommandType = CommandType.StoredProcedure
            GetPawnCreditInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Pawn_Credit
            .SelectCommand = GetPawnCreditInfoCommand
            .SelectCommand.Transaction = ts
            GetPawnCreditInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Pawn_Credit
            .Fill(tempDs, Table_Pawn_Credit)
        End With


        Return tempDs

    End Function


    Public Function UpdatePawnCredit(ByVal PawnCreditSet As DataSet)

        If PawnCreditSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If PawnCreditSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Pawn_Credit)

        With dsCommand_Pawn_Credit
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(PawnCreditSet, Table_Pawn_Credit)


        End With

        PawnCreditSet.AcceptChanges()

    End Function

    
    

End Class

