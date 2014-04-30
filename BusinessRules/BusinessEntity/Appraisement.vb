Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Appraisement

    Public Const Table_Project_Appraisement As String = "project_appraisement"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Appraisement As SqlDataAdapter

    '�����ѯ����
    Private GetAppraisementInfoCommand As SqlCommand
    Private GetMaxAppraisementNumCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Appraisement = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetAppraisementInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetAppraisementInfo(ByVal strSQL_Condition_Appraisement As String) As DataSet

        Dim tempDs As New DataSet()

        If GetAppraisementInfoCommand Is Nothing Then

            GetAppraisementInfoCommand = New SqlCommand("GetAppraisementInfo", conn)
            GetAppraisementInfoCommand.CommandType = CommandType.StoredProcedure
            GetAppraisementInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Appraisement
            .SelectCommand = GetAppraisementInfoCommand
            .SelectCommand.Transaction = ts
            GetAppraisementInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Appraisement
            .Fill(tempDs, Table_Project_Appraisement)
        End With

        Return tempDs

    End Function

    '��ȡ������к�
    Public Function GetMaxAppraisementNum(ByVal projectID As String) As Integer

        If GetMaxAppraisementNumCommand Is Nothing Then

            GetMaxAppraisementNumCommand = New SqlCommand("GetMaxAppraisementNum", conn)
            GetMaxAppraisementNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxAppraisementNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxAppraisementNumCommand.Parameters.Add(New SqlParameter("@maxAppraisementNum", SqlDbType.Int))
            GetMaxAppraisementNumCommand.Parameters.Item("@maxAppraisementNum").Direction = ParameterDirection.Output
            GetMaxAppraisementNumCommand.Transaction = ts
        End If

        GetMaxAppraisementNumCommand.Parameters("@projectID").Value = projectID
        GetMaxAppraisementNumCommand.ExecuteNonQuery()
        GetMaxAppraisementNum = GetMaxAppraisementNumCommand.Parameters.Item("@maxAppraisementNum").Value
    End Function

    '������Ŀ������Ϣ
    Public Function UpdateAppraisement(ByVal AppraisementSet As DataSet)

        If AppraisementSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If AppraisementSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Appraisement)

        With dsCommand_Appraisement
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(AppraisementSet, Table_Project_Appraisement)

            AppraisementSet.AcceptChanges()
        End With


    End Function


End Class

