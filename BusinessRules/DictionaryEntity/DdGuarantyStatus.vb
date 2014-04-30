
Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class DdGuarantyStatus

    Public Const Table_DdGuarantyStatus As String = "dd_guaranty_status"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_DdGuarantyStatus As SqlDataAdapter

    '�����ѯ����
    Private GetDdGuarantyStatusInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_DdGuarantyStatus = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetDdGuarantyStatusInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetDdGuarantyStatusInfo(ByVal strSQL_Condition_DdGuarantyStatus As String) As DataSet

        Dim tempDs As New DataSet()

        If GetDdGuarantyStatusInfoCommand Is Nothing Then

            GetDdGuarantyStatusInfoCommand = New SqlCommand("GetDdGuarantyStatusInfo", conn)
            GetDdGuarantyStatusInfoCommand.CommandType = CommandType.StoredProcedure
            GetDdGuarantyStatusInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_DdGuarantyStatus
            .SelectCommand = GetDdGuarantyStatusInfoCommand
            .SelectCommand.Transaction = ts
            GetDdGuarantyStatusInfoCommand.Parameters("@Condition").Value = strSQL_Condition_DdGuarantyStatus
            .Fill(tempDs, Table_DdGuarantyStatus)
        End With

        Return tempDs

    End Function

    '������Ŀ������Ϣ
    Public Function UpdateDdGuarantyStatus(ByVal DdGuarantyStatusSet As DataSet)

        If DdGuarantyStatusSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If DdGuarantyStatusSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_DdGuarantyStatus)

        With dsCommand_DdGuarantyStatus
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(DdGuarantyStatusSet, Table_DdGuarantyStatus)

        End With

        DdGuarantyStatusSet.AcceptChanges()
    End Function
End Class
