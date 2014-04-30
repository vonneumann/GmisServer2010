Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class SignaturePlan
    Public Const Table_Signature_Plan As String = "signature_plan"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_SignaturePlan As SqlDataAdapter

    '�����ѯ����
    Private GetSignaturePlanInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_SignaturePlan = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetSignaturePlanInfo("null")
    End Sub

    '��ȡǩԼ�ƻ���Ϣ
    Public Function GetSignaturePlanInfo(ByVal strSQL_Condition_SignaturePlan As String) As DataSet

        Dim tempDs As New DataSet()

        If GetSignaturePlanInfoCommand Is Nothing Then

            GetSignaturePlanInfoCommand = New SqlCommand("GetSignaturePlanInfo", conn)
            GetSignaturePlanInfoCommand.CommandType = CommandType.StoredProcedure
            GetSignaturePlanInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_SignaturePlan
            .SelectCommand = GetSignaturePlanInfoCommand
            .SelectCommand.Transaction = ts
            GetSignaturePlanInfoCommand.Parameters("@Condition").Value = strSQL_Condition_SignaturePlan
            .Fill(tempDs, Table_Signature_Plan)
        End With

        Return tempDs

    End Function


    '����ǩԼ�ƻ���Ϣ
    Public Function UpdateSignaturePlan(ByVal SignaturePlanSet As DataSet)

        If SignaturePlanSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If SignaturePlanSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_SignaturePlan)

        With dsCommand_SignaturePlan
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(SignaturePlanSet, Table_Signature_Plan)

            SignaturePlanSet.AcceptChanges()
        End With


    End Function
End Class
