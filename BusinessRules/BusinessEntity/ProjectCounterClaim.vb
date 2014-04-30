Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectCounterClaim
    Private Const Table_Project_Counter As String = "Project_counterclaim"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectCounter As SqlDataAdapter

    '�����ѯ����
    Private GetProjectCounterInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectCounter = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectCounterClaimInfo("null")
    End Sub

    '��ȡ������Ϣ
    Public Function GetProjectCounterClaimInfo(ByVal strSQL_Condition_ProjectCounterClaim As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectCounterInfoCommand Is Nothing Then

            GetProjectCounterInfoCommand = New SqlCommand("GetProjectCounterClaim", conn)
            GetProjectCounterInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectCounterInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectCounter
            .SelectCommand = GetProjectCounterInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectCounterInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectCounterClaim
            .Fill(tempDs, Table_Project_Counter)
        End With

        Return tempDs

    End Function

    '����������Ϣ
    Public Function UpdateProjectCounterClaim(ByVal ProjectCounterClaimSet As DataSet)


        If ProjectCounterClaimSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectCounterClaimSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectCounter)

        With dsCommand_ProjectCounter
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectCounterClaimSet, Table_Project_Counter)

            ProjectCounterClaimSet.AcceptChanges()
        End With


    End Function

End Class
