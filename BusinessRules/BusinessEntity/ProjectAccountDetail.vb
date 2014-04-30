Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ProjectAccountDetail

    Public Const Table_Project_Account_Detail As String = "project_account_detail"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_ProjectAccountDetail As SqlDataAdapter

    '�����ѯ����
    Private GetProjectAccountDetailInfoCommand As SqlCommand
    Private GetMaxProjectAccountDetailNumCommand As SqlCommand

    '��������
    Private ts As SqlTransaction


    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_ProjectAccountDetail = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetProjectAccountDetailInfo("null")
    End Sub

    '��ȡ��Ŀ������Ϣ
    Public Function GetProjectAccountDetailInfo(ByVal strSQL_Condition_ProjectAccountDetail As String) As DataSet

        Dim tempDs As New DataSet()

        If GetProjectAccountDetailInfoCommand Is Nothing Then

            GetProjectAccountDetailInfoCommand = New SqlCommand("GetProjectAccountDetailInfo", conn)
            GetProjectAccountDetailInfoCommand.CommandType = CommandType.StoredProcedure
            GetProjectAccountDetailInfoCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_ProjectAccountDetail
            .SelectCommand = GetProjectAccountDetailInfoCommand
            .SelectCommand.Transaction = ts
            GetProjectAccountDetailInfoCommand.Parameters("@Condition").Value = strSQL_Condition_ProjectAccountDetail
            .Fill(tempDs, Table_Project_Account_Detail)
        End With

        Return tempDs

    End Function


    '��ȡ������к�
    Public Function GetMaxProjectAccountDetailNum(ByVal projectID As String) As Integer

        If GetMaxProjectAccountDetailNumCommand Is Nothing Then

            GetMaxProjectAccountDetailNumCommand = New SqlCommand("GetMaxProjectAccountDetailNum", conn)
            GetMaxProjectAccountDetailNumCommand.CommandType = CommandType.StoredProcedure
            GetMaxProjectAccountDetailNumCommand.Parameters.Add(New SqlParameter("@projectID", SqlDbType.NVarChar))
            GetMaxProjectAccountDetailNumCommand.Parameters.Add(New SqlParameter("@maxProjectAccountDetailNum", SqlDbType.Int))
            GetMaxProjectAccountDetailNumCommand.Parameters.Item("@maxProjectAccountDetailNum").Direction = ParameterDirection.Output
            GetMaxProjectAccountDetailNumCommand.Transaction = ts
        End If

        GetMaxProjectAccountDetailNumCommand.Parameters("@projectID").Value = projectID
        GetMaxProjectAccountDetailNumCommand.ExecuteNonQuery()
        GetMaxProjectAccountDetailNum = GetMaxProjectAccountDetailNumCommand.Parameters.Item("@maxProjectAccountDetailNum").Value
    End Function

    '������Ŀ������Ϣ
    Public Function UpdateProjectAccountDetail(ByVal ProjectAccountDetailSet As DataSet)

        If ProjectAccountDetailSet Is Nothing Then
            Exit Function
        End If

        '�����¼��δ�����κα仯�����˳�����
        If ProjectAccountDetailSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_ProjectAccountDetail)

        With dsCommand_ProjectAccountDetail
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(ProjectAccountDetailSet, Table_Project_Account_Detail)

            ProjectAccountDetailSet.AcceptChanges()
        End With


    End Function
End Class
