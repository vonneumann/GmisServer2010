Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class Holiday
    Public Const Table_Holiday As String = "holiday"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_Holiday As SqlDataAdapter

    '�����ѯ����
    Private GetHolidayInfoCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_Holiday = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetHolidayInfo("null")
    End Sub

    '��ȡ������Ϣ
    Public Function GetHolidayInfo(ByVal strSQL_Condition_Holiday As String) As DataSet

        Dim tempDs As New DataSet()

        If GetHolidayInfoCommand Is Nothing Then

            GetHolidayInfoCommand = New SqlCommand("GetHolidayInfo", conn)
            GetHolidayInfoCommand.CommandType = CommandType.StoredProcedure
            GetHolidayInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_Holiday
            .SelectCommand = GetHolidayInfoCommand
            .SelectCommand.Transaction = ts
            GetHolidayInfoCommand.Parameters("@Condition").Value = strSQL_Condition_Holiday
            .Fill(tempDs, Table_Holiday)
        End With

        Return tempDs

    End Function

    '���¼�����Ϣ
    Public Function UpdateHoliday(ByVal HolidaySet As DataSet)

        If HolidaySet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If HolidaySet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_Holiday)

        With dsCommand_Holiday
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(HolidaySet, Table_Holiday)

        End With

        HolidaySet.AcceptChanges()
    End Function

End Class
