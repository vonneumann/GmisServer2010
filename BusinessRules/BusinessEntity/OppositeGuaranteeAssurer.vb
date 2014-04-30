Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class OppositeGuaranteeAssurer
    Public Const Table_OppositeGuaranteeAssurer As String = "opposite_guarantee_assurer"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_OppositeGuaranteeAssurer As SqlDataAdapter

    '�����ѯ����
    Private GetOppositeGuaranteeAssurerInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_OppositeGuaranteeAssurer = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetOppositeGuaranteeAssurerInfo("null")
    End Sub

    '��ȡ��������Ϣ
    Public Function GetOppositeGuaranteeAssurerInfo(ByVal strSQL_Condition_OppositeGuaranteeAssurer As String) As DataSet

        Dim tempDs As New DataSet()

        If GetOppositeGuaranteeAssurerInfoCommand Is Nothing Then

            GetOppositeGuaranteeAssurerInfoCommand = New SqlCommand("GetOppositeGuaranteeAssurerInfo", conn)
            GetOppositeGuaranteeAssurerInfoCommand.CommandType = CommandType.StoredProcedure
            GetOppositeGuaranteeAssurerInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_OppositeGuaranteeAssurer
            .SelectCommand = GetOppositeGuaranteeAssurerInfoCommand
            .SelectCommand.Transaction = ts
            GetOppositeGuaranteeAssurerInfoCommand.Parameters("@Condition").Value = strSQL_Condition_OppositeGuaranteeAssurer
            .Fill(tempDs, Table_OppositeGuaranteeAssurer)
        End With

        Return tempDs

    End Function

    '������������Ϣ
    Public Function UpdateOppositeGuaranteeAssurer(ByVal OppositeGuaranteeAssurerSet As DataSet)

        If OppositeGuaranteeAssurerSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If OppositeGuaranteeAssurerSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_OppositeGuaranteeAssurer)

        With dsCommand_OppositeGuaranteeAssurer
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(OppositeGuaranteeAssurerSet, Table_OppositeGuaranteeAssurer)

        End With

        OppositeGuaranteeAssurerSet.AcceptChanges()

    End Function
End Class
