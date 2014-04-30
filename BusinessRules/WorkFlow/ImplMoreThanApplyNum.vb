Option Explicit On 

Imports System
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient

Public Class ImplMoreThanApplyNum
    Implements ICondition

    '�����������
    Private applyNum As Integer

    '����ȫ�����ݿ����Ӷ���
    Private conn As SqlConnection

    '��������
    Private ts As SqlTransaction


    Public Sub New(ByVal DbConnection As SqlConnection, ByRef trans As SqlTransaction)
        MyBase.New()
        conn = DbConnection


        '�����ݿ�����
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        '�����ⲿ����
        ts = trans

    End Sub


    Public Function GetResult(ByVal workFlowID As String, ByVal projectID As String, ByVal taskID As String, ByVal expFlag As String, ByVal transCondition As String) As Boolean Implements ICondition.GetResult

        '��ͬ�����������>�����������
        If expFlag = "��ͬ��" Then

            '��ȡ�������
            Dim i As Integer = GetApplyNum(ProjectID)

            '�ж���������Ƿ�>�����������
            If i > _ApplyNumLimit Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    '��ȡ�������
    Private Function GetApplyNum(ByVal ProjectID As String) As Integer
        Dim IntentLetter As New IntentLetter(conn, ts)
        Dim strSql As String = "{project_code=" & "'" & ProjectID & "'" & "}"
        Dim dsTemp As DataSet = IntentLetter.GetIntentLetterInfo(strSql)

        '��¼������������Ĵ���
        applyNum = dsTemp.Tables(0).Rows.Count

    End Function
End Class
