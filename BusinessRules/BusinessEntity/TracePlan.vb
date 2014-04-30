Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class TracePlan
    Public Const Table_Trace_Plan As String = "trace_plan"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_TracePlan As SqlDataAdapter

    '�����ѯ����
    Private GetTracePlanInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_TracePlan = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetTracePlanInfo("null")
    End Sub

    '��ȡ����¼������Ϣ
    Public Function GetTracePlanInfo(ByVal strSQL_Condition_TracePlan As String) As DataSet

        Dim tempDs As New DataSet()

        If GetTracePlanInfoCommand Is Nothing Then

            GetTracePlanInfoCommand = New SqlCommand("GetTracePlanInfo", conn)
            GetTracePlanInfoCommand.CommandType = CommandType.StoredProcedure
            GetTracePlanInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_TracePlan
            .SelectCommand = GetTracePlanInfoCommand
            .SelectCommand.Transaction = ts
            GetTracePlanInfoCommand.Parameters("@Condition").Value = strSQL_Condition_TracePlan
            .Fill(tempDs, Table_Trace_Plan)
        End With

        Return tempDs

    End Function

    '���¼���¼������Ϣ
    Public Function UpdateTracePlan(ByVal TracePlanSet As DataSet)

        If TracePlanSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If TracePlanSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_TracePlan)

        With dsCommand_TracePlan
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(TracePlanSet, Table_Trace_Plan)

        End With

        TracePlanSet.AcceptChanges()

    End Function
End Class
