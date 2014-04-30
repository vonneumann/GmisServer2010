Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class UserPost
    Private Const Table_UserPost As String = "TUserPost"

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '����ȫ�����ݿ�����������
    Private dsCommand_UserPost As SqlDataAdapter

    '�����ѯ����
    Private GetUserPostCommand As SqlCommand

    '��������
    Private ts As SqlTransaction

    '���캯��
    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        'ʵ����������
        dsCommand_UserPost = New SqlDataAdapter()

        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

        '���������
        GetUserPostInfo("null")
    End Sub

    '��ȡ������Ϣ
    Public Function GetUserPostInfo(ByVal strSQL_Condition_UserPost As String) As DataSet

        Dim tempDs As New DataSet()

        If GetUserPostCommand Is Nothing Then

            GetUserPostCommand = New SqlCommand("dbo.GetUserPostInfo", conn)
            GetUserPostCommand.CommandType = CommandType.StoredProcedure
            GetUserPostCommand.Parameters.Add(New SqlParameter("@Condition", SqlDbType.NVarChar))

        End If

        With dsCommand_UserPost
            .SelectCommand = GetUserPostCommand
            .SelectCommand.Transaction = ts
            GetUserPostCommand.Parameters("@Condition").Value = strSQL_Condition_UserPost
            .Fill(tempDs, Table_UserPost)
        End With

        Return tempDs

    End Function

    '���¼�����Ϣ
    Public Function UpdateUserPost(ByVal UserPostSet As DataSet)

        If UserPostSet Is Nothing Then
            Exit Function
        End If


        '�����¼��δ�����κα仯�����˳�����
        If UserPostSet.HasChanges = False Then
            Exit Function
        End If

        Dim bd As SqlCommandBuilder = New SqlCommandBuilder(dsCommand_UserPost)

        With dsCommand_UserPost
            .InsertCommand = bd.GetInsertCommand
            .UpdateCommand = bd.GetUpdateCommand
            .DeleteCommand = bd.GetDeleteCommand

            .InsertCommand.Transaction = ts
            .UpdateCommand.Transaction = ts
            .DeleteCommand.Transaction = ts

            .Update(UserPostSet, Table_UserPost)

        End With

        UserPostSet.AcceptChanges()
    End Function
End Class
